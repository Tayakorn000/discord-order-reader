/**
 * Google Sheets integration for tracking parcel status.
 * Each day is organized into its own tab (e.g., "12/4/69").
 * Column Structure: A=Tracking Number, B=Check-in, C=Match Status, D=Timestamp
 */

const { google } = require('googleapis')
const fs = require('fs')
const path = require('path')

/**
 * Workaround for gaxios in pkg environment.
 * Forces gaxios to use the CJS version of node-fetch for compatibility.
 */
try {
  const gaxios = require('gaxios')
  const nodeFetch = require('node-fetch')
  const fetch = nodeFetch.default || nodeFetch
  gaxios.instance.defaults.fetchImplementation = fetch
} catch (err) {
  // Silent fail if gaxios or node-fetch is not available
}

const SHEET_ID = process.env.GOOGLE_SHEET_ID

/**
 * Initializes Google Auth using service account credentials.
 * Prioritizes GOOGLE_SERVICE_ACCOUNT environment variable over local credentials.json.
 * @returns {google.auth.GoogleAuth}
 */
function getAuth() {
  let credentials
  if (process.env.GOOGLE_SERVICE_ACCOUNT) {
    credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT)
  } else {
    const localPath = path.join(__dirname, 'credentials.json')
    if (!fs.existsSync(localPath)) {
      throw new Error('Google credentials missing. Set GOOGLE_SERVICE_ACCOUNT or add credentials.json')
    }
    credentials = JSON.parse(fs.readFileSync(localPath, 'utf8'))
  }
  return new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  })
}

/**
 * Returns a Google Sheets API instance.
 */
function api() {
  return google.sheets({ version: 'v4', auth: getAuth() })
}

/**
 * Returns the tab name for today in Thai format (D/M/YY).
 * Note: Year is converted to Buddhist Era (BE) shortened to 2 digits.
 * @returns {string}
 */
function todayTab() {
  const d = new Date()
  const thaiYear = (d.getFullYear() + 543) % 100
  return `${d.getDate()}/${d.getMonth() + 1}/${thaiYear}`
}

/**
 * Formats ISO string to Thai readable time format.
 * @param {string} isoStr 
 * @returns {string}
 */
function thaiTime(isoStr) {
  if (!isoStr) return ''
  const d = new Date(isoStr)
  const hh = d.getHours().toString().padStart(2, '0')
  const mm = d.getMinutes().toString().padStart(2, '0')
  const thaiYear = (d.getFullYear() + 543) % 100
  return `${hh}.${mm} น. (${d.getDate()}/${d.getMonth() + 1}/${thaiYear})`
}

/**
 * Retrieves metadata for all sheets in the spreadsheet.
 * @returns {Promise<Array<{title: string, sheetId: number}>>}
 */
async function getAllSheetMeta() {
  const info = await api().spreadsheets.get({ spreadsheetId: SHEET_ID })
  return info.data.sheets.map(s => ({
    title: s.properties.title,
    sheetId: s.properties.sheetId
  }))
}

/**
 * Finds an existing tab or creates a new one with headers and formatting.
 * @param {string} tabName 
 * @returns {Promise<number>} sheetId
 */
async function getOrCreateTab(tabName) {
  const sheets = await getAllSheetMeta()
  const found = sheets.find(s => s.title === tabName)
  if (found) return found.sheetId

  // Create new tab if not found
  const res = await api().spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: {
      requests: [{ addSheet: { properties: { title: tabName } } }],
    },
  })
  const newSheetId = res.data.replies[0].addSheet.properties.sheetId

  // Initialize header row
  await api().spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `'${tabName}'!A1:D1`,
    valueInputOption: 'RAW',
    requestBody: { values: [['หมายเลขพัสดุ', 'ตรวจรับ', 'จับคู่', 'เวลา']] },
  })

  // Apply header styling: Bold text, Gray background, Centered alignment
  await api().spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: {
      requests: [{
        repeatCell: {
          range: { sheetId: newSheetId, startRowIndex: 0, endRowIndex: 1 },
          cell: {
            userEnteredFormat: {
              backgroundColor: { red: 0.85, green: 0.85, blue: 0.85 },
              textFormat: { bold: true },
              horizontalAlignment: 'CENTER',
            }
          },
          fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
        },
      }],
    },
  })

  // Auto-resize columns to fit content
  await api().spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: {
      requests: [{
        autoResizeDimensions: {
          dimensions: { sheetId: newSheetId, dimension: 'COLUMNS', startIndex: 0, endIndex: 4 },
        }
      }],
    },
  })

  return newSheetId
}

/**
 * Appends new tracking numbers to today's tab.
 * Skips duplicates already present in the sheet.
 * @param {string[]} trackingNumbers 
 */
async function writeParcels(trackingNumbers) {
  if (!SHEET_ID) throw new Error('GOOGLE_SHEET_ID not set')
  const tab = todayTab()
  const sheetId = await getOrCreateTab(tab)

  const existing = await readTab(tab)
  const existSet = new Set(existing.map(r => r.tn))
  const newRows = trackingNumbers
    .filter(tn => !existSet.has(tn))
    .map(tn => [tn, '', 'ยังไม่ได้รับ', ''])

  if (!newRows.length) return { added: 0, tab }

  await api().spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: `'${tab}'!A:D`,
    valueInputOption: 'RAW',
    requestBody: { values: newRows },
  })

  // Highlight new rows as pending (Red)
  const startRow = existing.length + 1
  await colorRows(sheetId, startRow, startRow + newRows.length, false)

  return { added: newRows.length, tab }
}

/**
 * Reads data from a specific tab.
 * @param {string} tabName 
 * @returns {Promise<Array<Object>>}
 */
async function readTab(tabName) {
  const res = await api().spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `'${tabName}'!A:D`,
  })
  const rows = res.data.values ?? []
  return rows.slice(1) // Skip header
    .map((row, i) => ({
      tn: (row[0] ?? '').trim(),
      checked: (row[1] ?? '').trim(),
      status: (row[2] ?? '').trim(),
      time: (row[3] ?? '').trim(),
      rowIndex: i + 2, // 1-indexed, header is row 1
    }))
    .filter(r => r.tn)
}

/**
 * Aggregates all parcels from all tabs for cross-day lookups.
 * @returns {Promise<Object>}
 */
async function readAllParcels() {
  const sheets = await getAllSheetMeta()
  const result = {}
  for (const { title: tab } of sheets) {
    try {
      const rows = await readTab(tab)
      for (const row of rows) {
        result[row.tn] = {
          tab,
          rowIndex: row.rowIndex,
          received: row.status === 'รับแล้ว',
          time: row.time
        }
      }
    } catch (err) {
      // Ignore tabs that don't match data format
    }
  }
  return result
}

/**
 * Locates a parcel by tracking number, searching from newest tabs to oldest.
 * @param {string} trackingNumber 
 * @returns {Promise<Object|null>}
 */
async function lookupParcel(trackingNumber) {
  const sheets = await getAllSheetMeta()
  for (const { title: tab, sheetId } of sheets.reverse()) {
    try {
      const rows = await readTab(tab)
      const row = rows.find(r => r.tn === trackingNumber)
      if (row) return { ...row, tab, sheetId }
    } catch (err) {
      // Skip invalid tabs
    }
  }
  return null
}

/**
 * Updates a parcel status to "Received" and sets the current timestamp.
 * @param {string} trackingNumber 
 * @param {string} isoTime 
 */
async function markReceived(trackingNumber, isoTime) {
  if (!SHEET_ID) throw new Error('GOOGLE_SHEET_ID not set')
  const found = await lookupParcel(trackingNumber)
  if (!found) return false

  const { tab, rowIndex, sheetId } = found
  const timeStr = thaiTime(isoTime)

  // Update tracking status and timestamp in Sheet
  await api().spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `'${tab}'!B${rowIndex}:D${rowIndex}`,
    valueInputOption: 'RAW',
    requestBody: { values: [[trackingNumber, 'รับแล้ว', timeStr]] },
  })

  // Highlight row as completed (Green)
  await colorRows(sheetId, rowIndex - 1, rowIndex, true)
  return { tab, rowIndex, time: timeStr }
}

/**
 * Helper to update row background colors.
 * @param {number} sheetId 
 * @param {number} startRow 
 * @param {number} endRow 
 * @param {boolean} received - If true, uses green; otherwise red.
 */
async function colorRows(sheetId, startRow, endRow, received) {
  const bg = received
    ? { red: 0.565, green: 0.816, blue: 0.314 } // #90D050 Green
    : { red: 1.0, green: 0.2, blue: 0.2 }     // Red

  await api().spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: {
      requests: [{
        repeatCell: {
          range: { sheetId, startRowIndex: startRow, endRowIndex: endRow, startColumnIndex: 0, endColumnIndex: 4 },
          cell: { userEnteredFormat: { backgroundColor: bg } },
          fields: 'userEnteredFormat.backgroundColor',
        },
      }],
    },
  })
}

/**
 * Fetches today's data for local export (PDF/Excel generation).
 * @returns {Promise<Object|null>}
 */
async function readTodayForExport() {
  if (!SHEET_ID) return null
  try {
    const tab = todayTab()
    const rows = await readTab(tab)
    const parcels = {}
    for (const r of rows) {
      parcels[r.tn] = { received: r.status === 'รับแล้ว', receivedAt: r.time, date: tab }
    }
    return { parcels }
  } catch (err) {
    return null
  }
}

/**
 * Clears today's data rows while preserving the header.
 * @returns {Promise<Object>} Summary of cleared items.
 */
async function clearToday() {
  const tab = todayTab()
  const sheets = await getAllSheetMeta()
  const found = sheets.find(s => s.title === tab)
  if (!found) return { tab, cleared: 0 }

  const rows = await readTab(tab)
  if (!rows.length) return { tab, cleared: 0 }

  // Clear all data rows below header
  await api().spreadsheets.values.clear({
    spreadsheetId: SHEET_ID,
    range: `'${tab}'!A2:Z9999`,
  })

  // Reset background color to white
  await api().spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: {
      requests: [{
        repeatCell: {
          range: { sheetId: found.sheetId, startRowIndex: 1, endRowIndex: rows.length + 1 },
          cell: { userEnteredFormat: { backgroundColor: { red: 1, green: 1, blue: 1 } } },
          fields: 'userEnteredFormat.backgroundColor',
        },
      }],
    },
  })

  return { tab, cleared: rows.length }
}

module.exports = {
  writeParcels,
  lookupParcel,
  markReceived,
  readTodayForExport,
  clearToday,
  todayTab,
  SHEET_ID: () => SHEET_ID
}
