const fs     = require('fs')
const path   = require('path')

// ── Polyfill: pkg bundles Node.js with minimal ICU (utf-8 + utf-16le only) ──
// fontkit uses TextDecoder('ascii'/'windows-1252') which throws in pkg env.
// ASCII is a subset of UTF-8, so mapping to utf-8 is safe for font name parsing.
if (typeof TextDecoder !== 'undefined') {
  const _OrigDecoder = TextDecoder
  global.TextDecoder = class PatchedTextDecoder extends _OrigDecoder {
    constructor(encoding, options) {
      const enc = (encoding || 'utf-8').toLowerCase()
      // Pkg supports only utf-8 and utf-16le — map everything else to utf-8
      const SUPPORTED = new Set(['utf-8', 'utf8', 'utf-16le', 'utf-16', 'unicode-1-1-utf-8'])
      super(SUPPORTED.has(enc) ? enc : 'utf-8', options)
    }
  }
}

// ── Runtime directory: ใช้ directory ของ .exe เมื่อรันผ่าน pkg ──────────────
const RUNTIME_DIR = process.pkg ? path.dirname(process.execPath) : __dirname

// Load .env — ลองหลาย path เผื่อ Windows bat เปลี่ยน cwd
;(function loadEnv() {
  const dirs = [RUNTIME_DIR, process.cwd()]
  const names = ['.env', 'env']  // Windows บางครั้ง copy มาเป็น "env" ไม่มีจุด
  for (const dir of dirs) {
    for (const name of names) {
      const p = path.join(dir, name)
      if (fs.existsSync(p)) {
        require('dotenv').config({ path: p })
        console.log('Loaded env from:', p)
        return
      }
    }
  }
  console.warn('Warning: env file not found in', dirs)
})()

// ── Error logging — เขียน error ลง bot-error.log เสมอ ───────────────────────
const LOG_PATH = path.join(RUNTIME_DIR, 'bot-error.log')
function writeLog(msg) {
  const line = `[${new Date().toISOString()}] ${msg}\n`
  try { fs.appendFileSync(LOG_PATH, line) } catch {}
  console.error(msg)
}
process.on('uncaughtException', err => {
  writeLog('CRASH: ' + err.stack)
  console.error('\n=== BOT CRASHED ===')
  console.error(err.message)
  console.error('\nดู bot-error.log ในโฟลเดอร์เดียวกับ Discord Bot.exe')
  process.exitCode = 1
})
process.on('unhandledRejection', err => {
  writeLog('UNHANDLED: ' + (err?.stack || err))
})

const { Client, GatewayIntentBits, Events, AttachmentBuilder } = require('discord.js')
const PDFDocument = require('pdfkit')
const XLSX   = require('xlsx')
const Jimp   = require('jimp')
const jsQR   = require('jsqr')
const fetch  = require('node-fetch')
const Sheets = require('./sheets')

// ── Paths ─────────────────────────────────────────────────────────────────────
const DB_PATH     = path.join(RUNTIME_DIR, 'db.json')
const IMAGES_DIR  = path.join(RUNTIME_DIR, 'images')
const FONT_PATH   = path.join(RUNTIME_DIR, 'Sarabun-Regular.ttf')
const FONT_BOLD   = path.join(RUNTIME_DIR, 'Sarabun-Bold.ttf')

if (!fs.existsSync(IMAGES_DIR)) fs.mkdirSync(IMAGES_DIR, { recursive: true })

// Download Thai font if missing
async function ensureFont() {
  const downloads = [
    { path: FONT_PATH, url: 'https://github.com/google/fonts/raw/main/ofl/sarabun/Sarabun-Regular.ttf' },
    { path: FONT_BOLD, url: 'https://github.com/google/fonts/raw/main/ofl/sarabun/Sarabun-Bold.ttf'    },
  ]
  for (const { path: fp, url } of downloads) {
    if (!fs.existsSync(fp)) {
      try {
        console.log(`Downloading font: ${path.basename(fp)}…`)
        const res = await fetch(url)
        fs.writeFileSync(fp, await res.buffer())
      } catch (e) { console.error('Font download failed:', e.message) }
    }
  }
}

// ── DB ────────────────────────────────────────────────────────────────────────
function loadDb() {
  if (fs.existsSync(DB_PATH)) {
    try { return JSON.parse(fs.readFileSync(DB_PATH, 'utf8')) } catch {}
  }
  return { parcels: {} }
}
function saveDb(db) {
  try { fs.writeFileSync(DB_PATH, JSON.stringify(db, null, 2)) } catch {}
}

// ── Excel parser ──────────────────────────────────────────────────────────────
function parseExcel(buffer) {
  const wb  = XLSX.read(buffer, { type: 'buffer' })
  const ws  = wb.Sheets[wb.SheetNames[0]]
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' })

  let headerRow = -1, parcelCol = -1
  for (let i = 0; i < raw.length; i++) {
    const col = raw[i].findIndex(v => String(v).trim() === 'พัสดุ')
    if (col !== -1) { headerRow = i; parcelCol = col; break }
  }
  if (headerRow === -1) return { error: 'ไม่พบคอลัมน์ "พัสดุ" ในไฟล์' }

  const trackingNumbers = []
  for (let i = headerRow + 1; i < raw.length; i++) {
    const val = String(raw[i][parcelCol] ?? '').trim()
    if (val) trackingNumbers.push(val)
  }
  return { trackingNumbers }
}

// ── QR / Barcode scanner ──────────────────────────────────────────────────────
function safeCrop(img, x, y, w, h) {
  const W = img.bitmap.width, H = img.bitmap.height
  x = Math.max(0, Math.round(x)); y = Math.max(0, Math.round(y))
  w = Math.min(W - x, Math.round(w)); h = Math.min(H - y, Math.round(h))
  if (w < 60 || h < 60) return null
  return img.clone().crop(x, y, w, h)
}

// Normalize ภาพให้ไม่เกิน 1200px ก่อนทำอะไร — ลด RAM และความเร็ว
function normalize(img) {
  const W = img.bitmap.width
  if (W > 1400) return img.clone().resize(1200, Jimp.AUTO)
  if (W < 400)  return img.clone().resize(800, Jimp.AUTO)
  return img.clone()
}

// Fast-path crops: 6 zones ที่สำคัญที่สุด
function getFastCrops(img) {
  const W = img.bitmap.width, H = img.bitmap.height
  const crops = [], add = (x,y,w,h) => { const c=safeCrop(img,x,y,w,h); if(c) crops.push(c) }
  add(0,       0,       W,       H)        // ทั้งรูป
  add(0,       0,       W,       H*0.5)    // ครึ่งบน
  add(0,       H*0.5,   W,       H*0.5)    // ครึ่งล่าง
  add(W*0.1,   H*0.1,   W*0.8,   H*0.8)   // center 80%
  add(0,       H*0.25,  W,       H*0.5)    // แถบกลาง
  add(W*0.25,  H*0.25,  W*0.5,   H*0.5)   // center 50%
  return crops
}

// Deep-path crops: เพิ่ม quadrant และ strips สำหรับรูปที่อ่านยาก
function getDeepCrops(img) {
  const W = img.bitmap.width, H = img.bitmap.height
  const crops = [], add = (x,y,w,h) => { const c=safeCrop(img,x,y,w,h); if(c) crops.push(c) }
  add(0,       0,       W*0.5,   H*0.5)
  add(W*0.5,   0,       W*0.5,   H*0.5)
  add(0,       H*0.5,   W*0.5,   H*0.5)
  add(W*0.5,   H*0.5,   W*0.5,   H*0.5)
  add(0,       H*0.15,  W,       H*0.35)
  add(0,       H*0.5,   W,       H*0.35)
  return crops
}

function decodeQR(img, inv) {
  try {
    if (!img?.bitmap?.width || !img?.bitmap?.height || !img?.bitmap?.data) return null
    const { data, width, height } = img.bitmap
    const r = jsQR(data, width, height, { inversionAttempts: inv })
    return r ? r.data.trim() : null
  } catch { return null }
}

// ลอง decode QR จาก 1 รูป ด้วย 4 มุมหลัก
function tryQR(v) {
  for (const inv of ['dontInvert', 'onlyInvert']) {
    for (const angle of [0, 90, 180, 270]) {
      const rotated = angle === 0 ? v : v.clone().rotate(angle)
      const r = decodeQR(rotated, inv)
      if (r) return r
    }
  }
  return null
}

// ลอง decode QR พร้อม diagonal angles (ช้ากว่า ใช้เฉพาะ deep pass)
function tryQRDeep(v) {
  for (const inv of ['dontInvert', 'onlyInvert']) {
    for (const angle of [0, 45, 90, 135, 180, 225, 270, 315]) {
      const rotated = angle === 0 ? v : v.clone().rotate(angle)
      const r = decodeQR(rotated, inv)
      if (r) return r
    }
  }
  return null
}

async function scanQR(base) {
  // FAST PASS: ภาพปกติ + grayscale contrast, 4 มุม
  for (const crop of getFastCrops(base)) {
    const variants = [
      crop,
      crop.clone().grayscale().contrast(0.5),
      crop.clone().grayscale().contrast(0.8),
    ]
    for (const v of variants) {
      const r = tryQR(v); if (r) return r
    }
  }
  // DEEP PASS: เพิ่ม crop + sharpening + 8 มุม (เฉพาะรูปยาก)
  for (const crop of getDeepCrops(base)) {
    const variants = [
      crop.clone().grayscale().contrast(0.6),
      crop.clone().grayscale().convolute([[0,-1,0],[-1,5,-1],[0,-1,0]]),
      crop.clone().grayscale().contrast(0.7).convolute([[0,-1,0],[-1,5,-1],[0,-1,0]]),
    ]
    for (const v of variants) {
      const r = tryQRDeep(v); if (r) return r
    }
  }
  return null
}

let _readBarcodes = null
let _zxingUnavailable = false
async function getZxing() {
  if (_zxingUnavailable) return null
  if (_readBarcodes) return _readBarcodes
  try {
    const m = await import('zxing-wasm/reader')
    _readBarcodes = m.readBarcodes
  } catch {
    _zxingUnavailable = true
    return null
  }
  return _readBarcodes
}

async function scanBarcode(img) {
  const readBarcodes = await getZxing()
  if (!readBarcodes) return null
  const toImageData = v => ({
    data: new Uint8ClampedArray(v.bitmap.data.buffer),
    width: v.bitmap.width, height: v.bitmap.height,
  })
  try {
    for (const crop of [...getFastCrops(img), ...getDeepCrops(img)]) {
      for (const angle of [0, 90, 180, 270]) {
        try {
          const rotated = angle === 0 ? crop : crop.clone().rotate(angle)
          if (!rotated?.bitmap?.width) continue
          const results = await readBarcodes(toImageData(rotated), { tryHarder: true })
          const text = results?.[0]?.text?.trim()
          if (text) return text
        } catch { /* skip */ }
      }
    }
  } catch (err) { console.error('Barcode scan error:', err.message) }
  return null
}

async function scan(imageBuffer) {
  let img
  try { img = await Jimp.read(imageBuffer) } catch { return null }
  // normalize ภาพก่อนเสมอ — ลด RAM และเวลา
  const base = normalize(img)
  const qr = await scanQR(base)
  if (qr) return { text: qr, type: 'QR' }
  const bc = await scanBarcode(base)
  if (bc) return { text: bc, type: 'Barcode' }
  return null
}

// ── Save thumbnail for PDF ────────────────────────────────────────────────────
async function saveThumbnail(trackingNumber, imageBuffer) {
  try {
    const img  = await Jimp.read(imageBuffer)
    const dest = path.join(IMAGES_DIR, `${trackingNumber.replace(/[^a-zA-Z0-9]/g, '_')}.jpg`)
    await img.resize(300, Jimp.AUTO).quality(75).writeAsync(dest)
    return dest
  } catch { return null }
}

// ── PDF generator ─────────────────────────────────────────────────────────────
function thaiTime(isoString) {
  if (!isoString) return ''
  const d = new Date(isoString)
  const hh = d.getHours().toString().padStart(2, '0')
  const mm = d.getMinutes().toString().padStart(2, '0')
  const day   = d.getDate()
  const month = d.getMonth() + 1
  const year  = (d.getFullYear() + 543) % 100  // 2-digit BE year
  return `${hh}.${mm} น. (${day}/${month}/${year})`
}

async function generatePDF(db) {
  const font     = fs.existsSync(FONT_PATH) ? FONT_PATH : 'Helvetica'
  const fontBold = fs.existsSync(FONT_BOLD) ? FONT_BOLD : font

  const doc    = new PDFDocument({ size: 'A4', margin: 0, autoFirstPage: true })
  const chunks = []
  doc.on('data', c => chunks.push(c))

  const ML = 30, MT = 30
  const PAGE_W = 595, PAGE_H = 842
  const TABLE_W = PAGE_W - ML * 2

  // Column definitions (total = TABLE_W = 535)
  const COLS = [
    { label: 'หมายเลขพัสดุ', w: 148 },
    { label: 'ตรวจรับ',       w: 148 },
    { label: 'จับคู่',         w: 72  },
    { label: 'ภาพถ่าย',        w: 82  },
    { label: 'เวลา',           w: 85  },
  ]

  const HDR_H  = 28
  const ROW_H  = 24
  const IMG_H  = 72

  // Cell drawing helper
  function cell(x, y, w, h, { text = '', bg = null, fg = '#000000', align = 'center', fontSize = 7.5, imgPath = null, bold = false } = {}) {
    doc.save()
    if (bg) doc.rect(x, y, w, h).fill(bg)
    doc.rect(x, y, w, h).stroke('#AAAAAA')
    if (imgPath && fs.existsSync(imgPath)) {
      try { doc.image(imgPath, x + 2, y + 2, { fit: [w - 4, h - 4] }) } catch {}
    }
    if (text) {
      const f = bold ? fontBold : font
      doc.font(f).fontSize(fontSize).fillColor(fg)
        .text(text, x + 3, y + (h - fontSize * 1.3) / 2 + 1, { width: w - 6, align, lineBreak: false, ellipsis: true })
    }
    doc.restore()
  }

  function drawHeader(y) {
    let x = ML
    for (const col of COLS) {
      cell(x, y, col.w, HDR_H, { text: col.label, bg: '#D9D9D9', bold: true, fontSize: 8.5 })
      x += col.w
    }
    return y + HDR_H
  }

  let y = drawHeader(MT)
  const entries = Object.entries(db.parcels)

  for (const [tn, info] of entries) {
    const hasImg = info.imagePath && fs.existsSync(info.imagePath)
    const rowH   = hasImg ? IMG_H : ROW_H

    // Page break
    if (y + rowH > PAGE_H - MT - 90) {
      doc.addPage()
      y = drawHeader(MT)
    }

    const matchBg = info.received ? '#92D050' : '#FF0000'
    const matchFg = info.received ? '#000000' : '#FFFFFF'
    const matchTx = info.received ? 'สำเร็จ' : 'ไม่สำเร็จ'

    let x = ML
    cell(x, y, COLS[0].w, rowH, { text: tn, align: 'left' });                                            x += COLS[0].w
    cell(x, y, COLS[1].w, rowH, { text: info.received ? tn : '', align: 'left' });                        x += COLS[1].w
    cell(x, y, COLS[2].w, rowH, { text: matchTx, bg: matchBg, fg: matchFg, bold: info.received });        x += COLS[2].w
    cell(x, y, COLS[3].w, rowH, { imgPath: hasImg ? info.imagePath : null });                             x += COLS[3].w
    cell(x, y, COLS[4].w, rowH, { text: thaiTime(info.receivedAt) });

    y += rowH
  }

  // Footer
  y += 50
  doc.font(fontBold).fontSize(40).fillColor('#000000')
    .text('รบกวนพี่ๆขนส่งเซ็นต์รับด้วยนะคะ', ML, y, { width: TABLE_W, align: 'center' })
  y += 60
  doc.font(font).fontSize(40).fillColor('#000000')
    .text('(................................................)', ML, y, { width: TABLE_W, align: 'center' })

  doc.end()
  return new Promise(resolve => doc.on('end', () => resolve(Buffer.concat(chunks))))
}

// ── Excel summary ─────────────────────────────────────────────────────────────
function buildSummaryExcel(db) {
  const rows = Object.entries(db.parcels).map(([tn, info], i) => ({
    'ลำดับ': i + 1, 'พัสดุ': tn,
    'สถานะเข้ารับ': info.received ? 'รับแล้ว' : 'ยังไม่ได้รับ',
  }))
  const ws = XLSX.utils.json_to_sheet(rows, { header: ['ลำดับ', 'พัสดุ', 'สถานะเข้ารับ'] })
  ws['!cols'] = [{ wch: 8 }, { wch: 25 }, { wch: 16 }]
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, 'สรุปพัสดุ')
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' })
}

// ── Bot ───────────────────────────────────────────────────────────────────────
const { DISCORD_TOKEN, CHANNEL_ID } = process.env
if (!DISCORD_TOKEN) { console.error('❌  DISCORD_TOKEN missing'); process.exit(1) }
if (!CHANNEL_ID)    { console.error('❌  CHANNEL_ID missing');    process.exit(1) }

const client = new Client({
  intents: [GatewayIntentBits.Guilds, GatewayIntentBits.GuildMessages, GatewayIntentBits.MessageContent],
})

client.once(Events.ClientReady, () => {
  console.log(`✅  Logged in as ${client.user.tag}`)
  console.log(`📦  Watching channel ${CHANNEL_ID}`)
})

client.on(Events.MessageCreate, async (message) => {
  if (message.author.bot) return
  if (message.channelId !== CHANNEL_ID) return

  const text = message.content.trim()

  // ── pdf ─────────────────────────────────────────────────────────────────────
  if (text === 'pdf') {
    try {
      await message.react('⏳')
      // Prefer Sheets data; merge with local image paths from JSON
      let db = Sheets.SHEET_ID() ? (await Sheets.readTodayForExport()) : null
      if (!db) db = loadDb()
      // Merge imagePath from local db (Sheets mode stores images locally)
      if (Sheets.SHEET_ID()) {
        const localDb = loadDb()
        for (const [tn, info] of Object.entries(db.parcels)) {
          if (localDb.parcels[tn]?.imagePath) info.imagePath = localDb.parcels[tn].imagePath
        }
      }
      const total = Object.keys(db.parcels).length
      if (!total) { await message.reply('❌ ยังไม่มีข้อมูลพัสดุ'); return }

      const pdfBuf   = await generatePDF(db)
      const date     = new Date().toLocaleDateString('th-TH').replace(/\//g, '-')
      const received = Object.values(db.parcels).filter(v => v.received).length
      await message.reply({
        content: `📄 **รายงานพัสดุ**\n✅ รับแล้ว: **${received}** | ⏳ ค้างอยู่: **${total - received}**`,
        files: [new AttachmentBuilder(pdfBuf, { name: `รายงานพัสดุ_${date}.pdf` })],
      })
    } catch (err) {
      console.error('PDF error:', err)
      await message.reply('❌ เกิดข้อผิดพลาดในการสร้าง PDF')
    }
    return
  }

  // ── สรุป (Excel) ─────────────────────────────────────────────────────────────
  if (text === 'สรุป') {
    try {
      let db = Sheets.SHEET_ID() ? (await Sheets.readTodayForExport()) : null
      if (!db) db = loadDb()
      const total    = Object.keys(db.parcels).length
      const received = Object.values(db.parcels).filter(v => v.received).length
      if (!total) { await message.reply('❌ ยังไม่มีข้อมูลพัสดุ กรุณาอัปโหลดไฟล์ Excel ก่อน'); return }

      const buffer = buildSummaryExcel(db)
      const date   = new Date().toLocaleDateString('th-TH').replace(/\//g, '-')
      await message.reply({
        content: `📊 **สรุปพัสดุ**\n📦 ทั้งหมด: **${total}** | ✅ รับแล้ว: **${received}** | ⏳ ค้างอยู่: **${total - received}**`,
        files: [new AttachmentBuilder(buffer, { name: `สรุปพัสดุ_${date}.xlsx` })],
      })
    } catch (err) {
      console.error('สรุป error:', err)
      await message.reply('❌ เกิดข้อผิดพลาด')
    }
    return
  }

  // ── clear ────────────────────────────────────────────────────────────────────
  if (text === 'clear') {
    try {
      const today = Sheets.SHEET_ID() ? Sheets.todayTab() : new Date().toLocaleDateString('th-TH')

      // JSON db — clear today's parcels
      const db = loadDb()
      let cleared = 0
      for (const [tn, info] of Object.entries(db.parcels)) {
        if (info.date === today || info.date === new Date().toLocaleDateString('th-TH')) {
          delete db.parcels[tn]; cleared++
        }
      }
      saveDb(db)

      // Google Sheets — clear today's tab
      if (Sheets.SHEET_ID()) {
        const result = await Sheets.clearToday()
        cleared = result.cleared
        await message.reply(`🗑️ ล้างข้อมูลวันนี้แล้ว\n📄 Tab: **${result.tab}** | ลบ **${cleared}** รายการ`)
        return
      }

      await message.reply(`🗑️ ล้างข้อมูลวันนี้แล้ว **${cleared}** รายการ`)
    } catch (err) {
      console.error('Clear error:', err)
      await message.reply('❌ เกิดข้อผิดพลาดในการล้างข้อมูล')
    }
    return
  }

  // ── !status ──────────────────────────────────────────────────────────────────
  if (text === '!status') {
    const db       = loadDb()
    const total    = Object.keys(db.parcels).length
    const received = Object.values(db.parcels).filter(v => v.received).length
    await message.reply(`📊 ทั้งหมด: **${total}** | ✅ รับแล้ว: **${received}** | ⏳ ค้างอยู่: **${total - received}**`)
    return
  }

  if (!message.attachments.size) return

  // ── Excel upload ─────────────────────────────────────────────────────────────
  const xlsxFile = message.attachments.find(a => /\.(xlsx|xls)$/i.test(a.name ?? ''))
  if (xlsxFile) {
    try {
      const res    = await fetch(xlsxFile.url)
      const buffer = await res.buffer()
      const { trackingNumbers, error } = parseExcel(buffer)

      if (error) { await message.reply(`❌ ${error}`); return }
      if (!trackingNumbers.length) { await message.reply('❌ พบคอลัมน์ "พัสดุ" แต่ไม่มีข้อมูล'); return }

      // ── Google Sheets mode ──
      if (Sheets.SHEET_ID()) {
        try {
          await message.react('⏳')
          const { added, tab } = await Sheets.writeParcels(trackingNumbers)
          await message.reply(
            `✅ เขียนเข้า Google Sheet สำเร็จ\n` +
            `📄 Tab: **${tab}** | 📦 ทั้งหมด: **${trackingNumbers.length}** | เพิ่มใหม่: **${added}**`
          )
        } catch (e) {
          console.error('Sheets write error:', e.message)
          await message.reply(`❌ Google Sheets error: ${e.message}`)
        }
        return
      }

      // ── JSON db fallback ──
      const date = new Date().toLocaleDateString('th-TH')
      const db   = loadDb()
      let added  = 0
      for (const tn of trackingNumbers) {
        if (!db.parcels[tn]) { db.parcels[tn] = { date, received: false }; added++ }
      }
      saveDb(db)
      await message.reply(
        `✅ โหลดข้อมูลสำเร็จ\n📦 พบพัสดุ **${trackingNumbers.length}** รายการ (เพิ่มใหม่ **${added}**)\n📅 วันที่: ${date}`
      )
    } catch (err) {
      console.error('Excel error:', err)
      await message.reply('❌ เกิดข้อผิดพลาดในการอ่านไฟล์ Excel')
    }
    return
  }

  // ── Image scan ───────────────────────────────────────────────────────────────
  const imageFile = message.attachments.find(a => a.contentType?.startsWith('image/'))
  if (imageFile) {
    try { await message.react('🔍') } catch {}

    try {
      const res    = await fetch(imageFile.url)
      const buffer = await res.buffer()
      const found  = await scan(buffer)

      if (!found) {
        await message.reply('⚠️ ไม่พบ QR Code หรือ Barcode\nลองถ่ายให้ชัดขึ้น ตรงขึ้น หรือใกล้กว่านี้')
        return
      }

      const { text: code, type } = found
      const label = type === 'Barcode' ? '📊 Barcode' : '📱 QR Code'
      const now   = new Date().toISOString()

      // ── Google Sheets mode ──
      if (Sheets.SHEET_ID()) {
        try {
          const parcel = await Sheets.lookupParcel(code)
          if (!parcel) {
            await message.reply(`❌ ยังไม่เข้ารับ หรือ ไม่พบข้อมูลในระบบ\n${label}: \`${code}\``)
            return
          }
          if (parcel.received) {
            await message.reply(`✅ เข้ารับพัสดุแล้ว *(รับก่อนหน้านี้แล้ว)*\n${label}: \`${code}\`\n📄 Tab: ${parcel.tab}`)
            return
          }
          await Sheets.markReceived(code, now)
          // Save thumbnail locally for PDF
          const imgPath = await saveThumbnail(code, buffer)
          // Keep a local record for PDF image embedding
          const db = loadDb()
          if (!db.parcels[code]) db.parcels[code] = {}
          db.parcels[code].imagePath  = imgPath
          db.parcels[code].received   = true
          db.parcels[code].receivedAt = now
          saveDb(db)
          await message.reply(`✅ เข้ารับพัสดุแล้ว\n${label}: \`${code}\`\n📄 อัปเดต Sheet: **${parcel.tab}**`)
        } catch (e) {
          console.error('Sheets scan error:', e.message)
          await message.reply(`❌ Google Sheets error: ${e.message}`)
        }
        return
      }

      // ── JSON db fallback ──
      const db     = loadDb()
      const parcel = db.parcels[code]
      if (parcel) {
        if (parcel.received) {
          await message.reply(`✅ เข้ารับพัสดุแล้ว *(รับก่อนหน้านี้แล้ว)*\n${label}: \`${code}\``)
        } else {
          parcel.received   = true
          parcel.receivedAt = now
          parcel.imagePath  = await saveThumbnail(code, buffer) ?? null
          saveDb(db)
          await message.reply(`✅ เข้ารับพัสดุแล้ว\n${label}: \`${code}\``)
        }
      } else {
        await message.reply(`❌ ยังไม่เข้ารับ หรือ ไม่พบข้อมูลในระบบ\n${label}: \`${code}\``)
      }
    } catch (err) {
      console.error('Scan error:', err)
      await message.reply('❌ เกิดข้อผิดพลาดในการสแกนรูปภาพ')
    }
  }
})

// ── Start ─────────────────────────────────────────────────────────────────────
ensureFont().then(() => client.login(DISCORD_TOKEN))
