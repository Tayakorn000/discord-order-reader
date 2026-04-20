# Discord Order Reader

A Discord bot for a parcel intake workflow. Upload a daily manifest as Excel, photograph each arriving parcel so the bot decodes the tracking number from its QR / barcode, and generate the end-of-day PDF and Excel reports on demand.

Designed to run either as a **persistent Railway service** or as a **standalone Windows `.exe`** built with `pkg`.

---

## What it does

1. **Excel intake** — drop an `.xlsx` / `.xls` into the configured channel. The bot reads every sheet, finds the `พัสดุ` / `หมายเลขพัสดุ` column, deduplicates, and persists the tracking numbers to both `db.json` and (optionally) Google Sheets.
2. **QR / barcode scan on photo** — post a photo of an arriving parcel. The bot runs a multi-engine decode pipeline:
   - **jsQR** — fastest, handles clean QR codes
   - **ZXing (WASM)** — covers barcodes jsQR misses (Code 128, EAN, etc.)
   - **Google Vision API** — fallback OCR for damaged or low-contrast labels
   Each engine sees 14 preprocessed crops across 8 rotations with multiple binarization thresholds. The decoded tracking number is matched against today's manifest and marked as received.
3. **Report generation** — on-demand PDF (Thai font) and Excel summaries, plus manifest reset.

---

## Channel commands

| Command | Action |
|---|---|
| *(upload `.xlsx`)* | Parse manifest → store tracking numbers → reply with count |
| *(upload image)* | Decode QR/barcode → mark as received → reply with status |
| `pdf` | Generate today's PDF report with thumbnails |
| `สรุป` | Generate today's Excel summary |
| `clear` | Wipe today's data (local + Sheets tab) |
| `!status` | Show received / pending counts |

Only messages in the `CHANNEL_ID` channel are processed. All other channels are ignored.

---

## Tech stack

| Concern | Library |
|---|---|
| Discord client | `discord.js` v14 |
| Image preprocess | `sharp` (libvips) — 10× faster than Jimp |
| Primary QR decode | `jsqr` |
| Secondary decode | `zxing-wasm` |
| Tertiary decode | Google Vision (`googleapis`) |
| Excel I/O | `xlsx` (SheetJS) |
| PDF generation | `pdfkit` + Sarabun Thai font |
| Google Sheets | `googleapis` / `sheets.js` |
| Packaging (Windows) | `pkg` (Node 18, x64) |

---

## Configuration

All secrets are environment variables. Never commit them.

| Variable | Required | Purpose |
|---|---|---|
| `DISCORD_TOKEN` | ✅ | Discord bot token |
| `CHANNEL_ID` | ✅ | Channel the bot listens in |
| `GOOGLE_SHEET_ID` | optional | If set, parcels are mirrored to Google Sheets (one tab per day) |
| `GOOGLE_SERVICE_ACCOUNT` | optional | JSON string of a service-account key. If absent, the bot falls back to `credentials.json` on disk |
| `PORT` | optional | Port for the `/health` endpoint (default `3000`) |

Local `.env` example:

```dotenv
DISCORD_TOKEN=MTxxx.yyy.zzz
CHANNEL_ID=1234567890123456789
GOOGLE_SHEET_ID=1AbCdEf...xyz
GOOGLE_SERVICE_ACCOUNT={"type":"service_account","project_id":"…","private_key":"-----BEGIN…-----\n…","client_email":"…"}
```

---

## Running locally

```bash
npm install
cp .env.example .env        # then fill in the values
npm start
```

The bot logs in, logs `📦 Watching channel …`, and starts processing.

---

## Deploying to Railway

The repo includes `railway.json` and `Procfile`.

- The bot exposes `GET /health` on `$PORT`. While the Discord client is connecting it returns **503**; once ready it returns **200**.
- This drives zero-overlap deploys: Railway sends `SIGTERM` to the old instance only once the new one reports healthy. The bot handles `SIGTERM` gracefully — it stops accepting new events and waits for in-flight scans to finish.
- A three-layer deduplication scheme (reaction lock + in-memory `_seenMessages` TTL + message-ID check) prevents the triple-reply bug that arises when two instances briefly overlap.

---

## Windows `.exe` build

For on-premises PCs that should not run `node` directly:

```bash
npm run build:win
# → dist/Discord Bot.exe
```

The build uses `pkg` with `--no-bytecode` (required for `googleapis` compatibility). Several polyfills are baked into `index.js` to make `pkg` behave:

- `global.fetch` is injected from `node-fetch` v2 before `googleapis` loads, because gaxios uses dynamic `import('node-fetch')` which `pkg` cannot resolve.
- `TextDecoder` is patched to map unsupported encodings (`ascii`, `windows-1252`, …) to `utf-8`, because `pkg` bundles Node with minimal ICU. This keeps `fontkit` (inside `pdfkit`) working for Thai font parsing.
- `.env` is loaded relative to the `.exe` directory, not `process.cwd()`, so the shortcut / `.bat` launcher can live anywhere.
- A Windows-specific quirk: if antivirus strips the leading dot from `.env`, the loader also tries `env` (no dot).

Helper scripts for Windows operators:

- `install.bat` — installs Node.js dependencies on a fresh machine
- `เปิดบอท.bat` — starts the bot in a console window
- `build-app.sh` — macOS helper that assembles a `.app` bundle for testing

---

## Data model

### `db.json` (local fallback)

```json
{
  "parcels": {
    "TRACKING123": {
      "date": "13/4/2026",
      "received": true,
      "imagePath": "images/TRACKING123.jpg"
    }
  }
}
```

### Google Sheets layout

- One tab per day, named by day-of-month (e.g. `13469` for 13 Apr 2026 — matches the existing workflow used by the operator).
- Columns: tracking number, received flag, timestamp, image URL.
- `sheets.js` exposes `append`, `markReceived`, `clearToday`, and `readTodayForExport` — the last is used to generate the PDF / Excel from the authoritative Sheets data instead of the local JSON.

---

## Operational notes

- `images/`, `db.json`, `bot-error.log`, and the downloaded Thai fonts are all written to the **runtime directory** — i.e. the folder the bot was started from, or the folder containing the `.exe`. They are gitignored.
- On first run the bot downloads Sarabun Regular + Bold from the Google Fonts GitHub mirror if the font files are missing.
- Unhandled rejections and uncaught exceptions are appended to `bot-error.log`; the bot does **not** crash silently.

---

## License

Tayakorn Wetchakun
