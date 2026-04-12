/**
 * Google Sheets integration
 * Each day gets its own tab: "12/4/69", "13/4/69" …
 * Columns: A=หมายเลขพัสดุ  B=ตรวจรับ  C=จับคู่  D=เวลา
 */

const { google } = require('googleapis')
const fs   = require('fs')
const path = require('path')

// ── Force gaxios to use CJS node-fetch (pkg compat) ──────────────────────────
try {
  const gaxios  = require('gaxios')
  const nodeFetch = require('node-fetch')
  const fetch = nodeFetch.default || nodeFetch
  gaxios.instance.defaults.fetchImplementation = fetch
} catch {}

const SHEET_ID = process.env.GOOGLE_SHEET_ID

// ── Auth ──────────────────────────────────────────────────────────────────────
function getAuth() {
  let credentials
  if (process.env.GOOGLE_SERVICE_ACCOUNT) {
    credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT)
  } else {
    const localPath = path.join(__dirname, 'credentials.json')
    if (!fs.existsSync(localPath)) throw new Error('No Google credentials. Set GOOGLE_SERVICE_ACCOUNT env var or add credentials.json')
    credentials = JSON.parse(fs.readFileSync(localPath, 'utf8'))
  }
  return new google.auth.GoogleAuth({ credentials, scopes: ['https://www.googleapis.com/auth/spreadsheets'] })
}

function api() { return google.sheets({ version: 'v4', auth: getAuth() }) }

// ── Helpers ───────────────────────────────────────────────────────────────────
function todayTab() {
  const d = new Date()
  return `${d.getDate()}/${d.getMonth() + 1}/${(d.getFullYear() + 543) % 100}`
}

function thaiTime(isoStr) {
  if (!isoStr) return ''
  const d = new Date(isoStr)
  return `${d.getHours().toString().padStart(2,'0')}.${d.getMinutes().toString().padStart(2,'0')} น. (${d.getDate()}/${d.getMonth()+1}/${(d.getFullYear()+543)%100})`
}

// ── Sheet tab management ──────────────────────────────────────────────────────
async function getAllSheetMeta() {
  const info = await api().spreadsheets.get({ spreadsheetId: SHEET_ID })
  return info.data.sheets.map(s => ({ title: s.properties.title, sheetId: s.properties.sheetId }))
}

async function getOrCreateTab(tabName) {
  const sheets = await getAllSheetMeta()
  const found  = sheets.find(s => s.title === tabName)
  if (found) return found.sheetId

  // Create new tab
  const res = await api().spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: { requests: [{ addSheet: { properties: { title: tabName } } }] },
  })
  const newSheetId = res.data.replies[0].addSheet.properties.sheetId

  // Write header row
  await api().spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `'${tabName}'!A1:D1`,
    valueInputOption: 'RAW',
    requestBody: { values: [['หมายเลขพัสดุ', 'ตรวจรับ', 'จับคู่', 'เวลา']] },
  })

  // Format header: bold, gray background, centered
  await api().spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: { requests: [{
      repeatCell: {
        range: { sheetId: newSheetId, startRowIndex: 0, endRowIndex: 1 },
        cell: { userEnteredFormat: {
          backgroundColor: { red: 0.85, green: 0.85, blue: 0.85 },
          textFormat: { bold: true },
          horizontalAlignment: 'CENTER',
        }},
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
      },
    }]},
  })

  // Auto-resize columns
  await api().spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: { requests: [{ autoResizeDimensions: {
      dimensions: { sheetId: newSheetId, dimension: 'COLUMNS', startIndex: 0, endIndex: 4 },
    }}]},
  })

  return newSheetId
}

// ── Write tracking numbers to today's tab ────────────────────────────────────
async function writeParcels(trackingNumbers) {
  if (!SHEET_ID) throw new Error('GOOGLE_SHEET_ID not set')
  const tab     = todayTab()
  const sheetId = await getOrCreateTab(tab)

  // Read existing to avoid duplicates
  const existing = await readTab(tab)
  const existSet  = new Set(existing.map(r => r.tn))
  const newRows   = trackingNumbers.filter(tn => !existSet.has(tn)).map(tn => [tn, '', 'ยังไม่ได้รับ', ''])

  if (!newRows.length) return { added: 0, tab }

  await api().spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: `'${tab}'!A:D`,
    valueInputOption: 'RAW',
    requestBody: { values: newRows },
  })

  // Color new rows red (ยังไม่ได้รับ)
  const startRow = existing.length + 1  // 0-indexed, header = row 0
  await colorRows(sheetId, startRow, startRow + newRows.length, false)

  return { added: newRows.length, tab }
}

// ── Read a tab ────────────────────────────────────────────────────────────────
async function readTab(tabName) {
  const res = await api().spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: `'${tabName}'!A:D`,
  })
  const rows = res.data.values ?? []
  return rows.slice(1)  // skip header
    .map((row, i) => ({
      tn:         (row[0] ?? '').trim(),
      checked:    (row[1] ?? '').trim(),
      status:     (row[2] ?? '').trim(),
      time:       (row[3] ?? '').trim(),
      rowIndex:   i + 2,  // 1-indexed, row 1 = header
    }))
    .filter(r => r.tn)
}

// ── Read all parcels from all tabs (for cross-day lookup) ─────────────────────
async function readAllParcels() {
  const sheets  = await getAllSheetMeta()
  const result  = {}
  for (const { title: tab } of sheets) {
    try {
      const rows = await readTab(tab)
      for (const row of rows) {
        result[row.tn] = { tab, rowIndex: row.rowIndex, received: row.status === 'รับแล้ว', time: row.time }
      }
    } catch { /* skip non-data tabs */ }
  }
  return result
}

// ── Look up a single tracking number ─────────────────────────────────────────
async function lookupParcel(trackingNumber) {
  const sheets = await getAllSheetMeta()
  for (const { title: tab, sheetId } of sheets.reverse()) {  // newest tab first
    try {
      const rows = await readTab(tab)
      const row  = rows.find(r => r.tn === trackingNumber)
      if (row) return { ...row, tab, sheetId }
    } catch {}
  }
  return null
}

// ── Mark parcel as received, update sheet ─────────────────────────────────────
async function markReceived(trackingNumber, isoTime) {
  if (!SHEET_ID) throw new Error('GOOGLE_SHEET_ID not set')
  const found = await lookupParcel(trackingNumber)
  if (!found) return false

  const { tab, rowIndex, sheetId } = found
  const timeStr = thaiTime(isoTime)

  // Update B, C, D
  await api().spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `'${tab}'!B${rowIndex}:D${rowIndex}`,
    valueInputOption: 'RAW',
    requestBody: { values: [[trackingNumber, 'รับแล้ว', timeStr]] },
  })

  // Color row green
  await colorRows(sheetId, rowIndex - 1, rowIndex, true)
  return { tab, rowIndex, time: timeStr }
}

// ── Color helper — green = received, red = pending ────────────────────────────
async function colorRows(sheetId, startRow, endRow, received) {
  const bg = received
    ? { red: 0.565, green: 0.816, blue: 0.314 }   // #90D050 green
    : { red: 1.0,   green: 0.2,   blue: 0.2   }   // red

  await api().spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: { requests: [{
      repeatCell: {
        range: { sheetId, startRowIndex: startRow, endRowIndex: endRow, startColumnIndex: 0, endColumnIndex: 4 },
        cell: { userEnteredFormat: { backgroundColor: bg } },
        fields: 'userEnteredFormat.backgroundColor',
      },
    }]},
  })
}

// ── Read today's tab as db-compatible object (for PDF/Excel export) ───────────
async function readTodayForExport() {
  if (!SHEET_ID) return null
  try {
    const tab  = todayTab()
    const rows = await readTab(tab)
    const parcels = {}
    for (const r of rows) {
      parcels[r.tn] = { received: r.status === 'รับแล้ว', receivedAt: r.time, date: tab }
    }
    return { parcels }
  } catch { return null }
}

// ── Clear today's tab (delete all data rows, keep header) ────────────────────
async function clearToday() {
  const tab    = todayTab()
  const sheets = await getAllSheetMeta()
  const found  = sheets.find(s => s.title === tab)
  if (!found) return { tab, cleared: 0 }

  // Read to count rows
  const rows = await readTab(tab)
  if (!rows.length) return { tab, cleared: 0 }

  // Clear everything below header
  await api().spreadsheets.values.clear({
    spreadsheetId: SHEET_ID,
    range: `'${tab}'!A2:Z9999`,
  })

  // Remove background color on cleared rows
  await api().spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: { requests: [{
      repeatCell: {
        range: { sheetId: found.sheetId, startRowIndex: 1, endRowIndex: rows.length + 1 },
        cell: { userEnteredFormat: { backgroundColor: { red: 1, green: 1, blue: 1 } } },
        fields: 'userEnteredFormat.backgroundColor',
      },
    }]},
  })

  return { tab, cleared: rows.length }
}

module.exports = { writeParcels, lookupParcel, markReceived, readTodayForExport, clearToday, todayTab, SHEET_ID: () => SHEET_ID }
