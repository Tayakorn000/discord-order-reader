const fs     = require('fs')
const path   = require('path')

// ── Polyfill fetch: gaxios (googleapis dep) ทำ dynamic import('node-fetch') ──
// ── ใน pkg สิ่งนี้ fail — inject global.fetch จาก node-fetch v2 (CJS) แทน ──
if (typeof globalThis.fetch === 'undefined') {
  try {
    const nf = require('node-fetch')
    globalThis.fetch    = nf.default || nf
    globalThis.Headers  = nf.Headers
    globalThis.Request  = nf.Request
    globalThis.Response = nf.Response
  } catch {}
}

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
const sharp = require('sharp')

// ── Pre-warm zxing-wasm ───────────────────────────────────────────────────────
let _zxing = null
async function getZxing() {
  if (_zxing === false) return null
  if (_zxing) return _zxing
  try { const m = await import('zxing-wasm/reader'); _zxing = m.readBarcodes }
  catch { _zxing = false }
  return _zxing || null
}

// ── Decoders ──────────────────────────────────────────────────────────────────
async function decodeZxing(imageData) {
  const rb = await getZxing()
  if (!rb) return null
  try {
    const r = await rb(imageData, { tryHarder: true, formats: [] })
    return r?.[0]?.text?.trim() || null
  } catch { return null }
}

function decodeJsQR(imageData) {
  try {
    return jsQR(imageData.data, imageData.width, imageData.height, { inversionAttempts: 'attemptBoth' })?.data?.trim() || null
  } catch { return null }
}

async function decodeAny(imageData) {
  return (await decodeZxing(imageData)) || decodeJsQR(imageData) || null
}

// ── Sharp preprocessing → ImageData ──────────────────────────────────────────
// sharp runs natively (libvips) — เร็วกว่า jimp 10× และ threshold ตัด glare ได้จริง
async function sharpToImageData(input, options = {}) {
  try {
    const { width, threshold = false, sharpen = false } = options
    let p = sharp(input).rotate()  // auto-rotate by EXIF
    if (width)    p = p.resize(width, null, { fit: 'inside', withoutEnlargement: false })
    if (sharpen)  p = p.sharpen({ sigma: 1.5, m1: 1, m2: 2 })
    p = p.grayscale()
    if (threshold) p = p.threshold(128)  // binary: ตัด glare ออกเป็น pure black/white
    const { data, info } = await p.raw().toBuffer({ resolveWithObject: true })
    // grayscale raw = 1 channel → แปลงเป็น RGBA สำหรับ zxing/jsQR
    const rgba = Buffer.alloc(info.width * info.height * 4)
    for (let i = 0; i < info.width * info.height; i++) {
      const v = data[i]
      rgba[i*4] = v; rgba[i*4+1] = v; rgba[i*4+2] = v; rgba[i*4+3] = 255
    }
    return { data: new Uint8ClampedArray(rgba), width: info.width, height: info.height }
  } catch { return null }
}

// ── Crop buffer ด้วย sharp (ไม่ต้องโหลด jimp เลย) ───────────────────────────
async function sharpCrop(buffer, meta, x, y, w, h) {
  try {
    x = Math.max(0, Math.round(x)); y = Math.max(0, Math.round(y))
    w = Math.min(meta.width - x, Math.round(w))
    h = Math.min(meta.height - y, Math.round(h))
    if (w < 60 || h < 60) return null
    return await sharp(buffer).extract({ left: x, top: y, width: w, height: h }).toBuffer()
  } catch { return null }
}

// ── Google Vision fallback — locate label bounding box ───────────────────────
async function visionLocateLabel(imageBuffer) {
  try {
    const { google } = require('googleapis')
    const Sheets = require('./sheets')
    if (!Sheets.SHEET_ID()) return null  // ไม่มี credentials ให้ skip
    const auth = (() => {
      const credentials = process.env.GOOGLE_SERVICE_ACCOUNT
        ? JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT)
        : JSON.parse(fs.readFileSync(path.join(RUNTIME_DIR, 'credentials.json'), 'utf8'))
      return new google.auth.GoogleAuth({ credentials, scopes: ['https://www.googleapis.com/auth/cloud-platform'] })
    })()
    const vision = google.vision({ version: 'v1', auth })
    const base64 = imageBuffer.toString('base64')
    const res = await vision.images.annotate({
      requestBody: { requests: [{ image: { content: base64 }, features: [{ type: 'DOCUMENT_TEXT_DETECTION' }] }] }
    })
    const pages = res.data.responses?.[0]?.fullTextAnnotation?.pages
    if (!pages?.length) return null
    // หา bounding box ของ text block ทั้งหมด
    const verts = res.data.responses[0].textAnnotations?.[0]?.boundingPoly?.vertices
    if (!verts?.length) return null
    const xs = verts.map(v => v.x || 0), ys = verts.map(v => v.y || 0)
    return { x: Math.min(...xs), y: Math.min(...ys), w: Math.max(...xs) - Math.min(...xs), h: Math.max(...ys) - Math.min(...ys) }
  } catch { return null }
}

// ── Main scan pipeline ────────────────────────────────────────────────────────
async function scan(imageBuffer) {
  // อ่าน metadata ก่อน (ไม่โหลดทั้งรูปเข้า memory)
  let meta
  try { meta = await sharp(imageBuffer).rotate().metadata() } catch { return null }
  const W = meta.width || 1000, H = meta.height || 1000

  // normalize: cap ที่ 1400px, เพิ่มถ้าเล็กกว่า 500px
  const targetW = W > 1400 ? 1300 : W < 500 ? 800 : null
  const baseBuffer = targetW
    ? await sharp(imageBuffer).rotate().resize(targetW, null, { fit: 'inside' }).toBuffer()
    : await sharp(imageBuffer).rotate().toBuffer()
  const baseMeta = await sharp(baseBuffer).metadata()
  const BW = baseMeta.width, BH = baseMeta.height

  // ── STAGE 1: ทั้งรูป — 4 variants พร้อมกัน ────────────────────────────────
  const stage1 = await Promise.all([
    sharpToImageData(baseBuffer),                                         // raw
    sharpToImageData(baseBuffer, { sharpen: true }),                      // sharpen
    sharpToImageData(baseBuffer, { threshold: true }),                    // binary (glare killer)
    sharpToImageData(baseBuffer, { sharpen: true, threshold: true }),     // sharpen + binary
  ])
  for (const id of stage1) { if (!id) continue; const r = await decodeAny(id); if (r) return { text: r, type: 'Code' } }

  // ── STAGE 2: overlapping zones + zoom — concurrent ────────────────────────
  const zones = [
    [0,       0,       BW*0.6,  BH*0.6 ],
    [BW*0.4,  0,       BW*0.6,  BH*0.6 ],
    [0,       BH*0.4,  BW*0.6,  BH*0.6 ],
    [BW*0.4,  BH*0.4,  BW*0.6,  BH*0.6 ],
    [BW*0.1,  BH*0.1,  BW*0.8,  BH*0.8 ],
    [0,       BH*0.2,  BW,      BH*0.6 ],
  ]
  const stage2 = await Promise.all(zones.map(async ([x, y, w, h]) => {
    const crop = await sharpCrop(baseBuffer, baseMeta, x, y, w, h)
    if (!crop) return null
    // zoom 2× ให้ QR เล็กมี pixel พอ
    const id = await sharpToImageData(crop, { width: Math.min(Math.round(w)*2, 1200), sharpen: true, threshold: true })
    return id ? decodeAny(id) : null
  }))
  const found2 = stage2.find(r => r)
  if (found2) return { text: found2, type: 'Code' }

  // ── STAGE 3: 3×3 grid + zoom 3× — concurrent ─────────────────────────────
  const gw = Math.floor(BW/3), gh = Math.floor(BH/3)
  const cells = []
  for (let r = 0; r < 3; r++) for (let c = 0; c < 3; c++) cells.push([c*gw, r*gh, gw, gh])

  const stage3 = await Promise.all(cells.map(async ([x, y, w, h]) => {
    const crop = await sharpCrop(baseBuffer, baseMeta, x, y, w, h)
    if (!crop) return null
    const id = await sharpToImageData(crop, { width: Math.min(gw*3, 900), sharpen: true, threshold: true })
    return id ? decodeAny(id) : null
  }))
  const found3 = stage3.find(r => r)
  if (found3) return { text: found3, type: 'Code' }

  // ── STAGE 4 (fallback): Google Vision หา bbox ของฉลาก แล้ว crop + scan ───
  const bbox = await visionLocateLabel(imageBuffer)
  if (bbox) {
    const padding = 20
    const crop = await sharpCrop(baseBuffer, baseMeta,
      bbox.x - padding, bbox.y - padding,
      bbox.w + padding*2, bbox.h + padding*2)
    if (crop) {
      const id = await sharpToImageData(crop, { width: 1200, sharpen: true, threshold: true })
      if (id) { const r = await decodeAny(id); if (r) return { text: r, type: 'Code' } }
    }
  }

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

  // ── Image scan — fetch + scan ทุกรูปพร้อมกัน (Promise.all) ──────────────────
  const imageFiles = [...message.attachments.values()].filter(a => a.contentType?.startsWith('image/'))
  if (imageFiles.length === 0) return

  try { await message.react('🔍') } catch {}

  const now = new Date().toISOString()

  // ดาวน์โหลด + scan ทุกรูปพร้อมกัน
  const scanResults = await Promise.all(imageFiles.map(async (imageFile) => {
    try {
      const res    = await fetch(imageFile.url)
      const buffer = await res.buffer()
      const found  = await scan(buffer)
      return { imageFile, buffer, found }
    } catch (err) {
      console.error('Scan error:', err)
      return { imageFile, buffer: null, found: null }
    }
  }))

  const results = []
  const noRead  = []

  for (const { imageFile, buffer, found } of scanResults) {
    try {
      if (!found) { noRead.push(imageFile.name ?? 'ไม่ทราบชื่อ'); continue }

      const { text: code, type } = found

      // ── Google Sheets mode ──
      if (Sheets.SHEET_ID()) {
        try {
          const parcel = await Sheets.lookupParcel(code)
          if (!parcel) {
            results.push({ code, type, status: 'notfound' }); continue
          }
          if (parcel.received) {
            results.push({ code, type, status: 'already', tab: parcel.tab }); continue
          }
          await Sheets.markReceived(code, now)
          const imgPath = await saveThumbnail(code, buffer)
          const db = loadDb()
          if (!db.parcels[code]) db.parcels[code] = {}
          db.parcels[code].imagePath  = imgPath
          db.parcels[code].received   = true
          db.parcels[code].receivedAt = now
          saveDb(db)
          results.push({ code, type, status: 'ok', tab: parcel.tab })
        } catch (e) {
          console.error('Sheets scan error:', e.message)
          results.push({ code, type, status: 'error', err: e.message })
        }
        continue
      }

      // ── JSON db fallback ──
      const db     = loadDb()
      const parcel = db.parcels[code]
      if (!parcel) {
        results.push({ code, type, status: 'notfound' })
      } else if (parcel.received) {
        results.push({ code, type, status: 'already' })
      } else {
        parcel.received   = true
        parcel.receivedAt = now
        parcel.imagePath  = await saveThumbnail(code, buffer) ?? null
        saveDb(db)
        results.push({ code, type, status: 'ok' })
      }
    } catch (err) {
      console.error('Scan error:', err)
      noRead.push(imageFile.name ?? 'ไม่ทราบชื่อ')
    }
  }

  // ── สรุปผลทั้งหมดในข้อความเดียว ─────────────────────────────────────────
  const total = imageFiles.length
  const lines = []

  for (const r of results) {
    const label = r.type === 'Barcode' ? '📊' : '📱'
    if      (r.status === 'ok')       lines.push(`✅ ${label} \`${r.code}\`${r.tab ? ` — ${r.tab}` : ''}`)
    else if (r.status === 'already')  lines.push(`🔁 ${label} \`${r.code}\` *(รับแล้ว)*`)
    else if (r.status === 'notfound') lines.push(`❌ ${label} \`${r.code}\` — ไม่พบในระบบ`)
    else                              lines.push(`⚠️ ${label} \`${r.code}\` — ${r.err}`)
  }
  for (const name of noRead) lines.push(`⚠️ อ่านไม่ได้: ${name}`)

  const okCount = results.filter(r => r.status === 'ok').length
  const header  = `📦 สแกน **${total}** รูป | ✅ เข้ารับ **${okCount}** | ❌ อ่านไม่ได้ **${noRead.length}**`
  await message.reply([header, ...lines].join('\n'))
})

// ── Start ─────────────────────────────────────────────────────────────────────
ensureFont().then(() => {
  getZxing()  // pre-warm zxing-wasm ก่อน bot online เพื่อลด cold-start delay
  client.login(DISCORD_TOKEN)
})
