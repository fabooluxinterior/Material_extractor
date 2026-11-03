# Faboolux Materials Extractor (server-side, live sheet sync)

Upload an estimate Excel and get the materials back instantly. The server **only reads** your
database Google Sheet. It never writes to Google Sheets, and nothing is saved to disk.

## How the data flows
- The database sheet (`Cabinet_Data`, `Laminate_Data`, `masterdata`) is downloaded and kept in memory.
- It is re-checked every **60 s** in the background, so any edit you make in the sheet appears automatically.
- Uploads never wait for Google: they read the in-memory copy, so results are near-instant.
- If Google is unreachable, the last good copy keeps being used (no data loss, no downtime).
- Each upload also kicks off a background re-check if the last one is older than 10 s.

## Setup
1. In the database Google Sheet: **Share -> General access -> Anyone with the link -> Viewer**
   (needed so the server can read it without logins; the sheet stays read-only to others).
2. `npm install`
3. `npm start` and open http://localhost:3000

Optional environment variables:
| Variable | Default | Meaning |
|---|---|---|
| `DATABASE_SHEET_ID` | your database sheet ID | Sheet to read |
| `SYNC_INTERVAL_MS` | `60000` | Background sync interval |
| `REFRESH_TOKEN` | none | If set, `/api/refresh` needs `?token=...` |
| `DB_PATH` | none | Read a local .xlsx instead of Google (offline use) |
| `PORT` | `3000` | Server port |

## Truly instant updates (optional)
The "Sync now" button in the page forces an immediate re-read. To make the sheet push changes the
moment someone edits it, add this to the **database sheet's** Apps Script
(Extensions -> Apps Script), then add an installable trigger: Triggers -> Add -> `onSheetEdit` -> From spreadsheet -> On edit.
It only pings your server; it reads and writes nothing.

```js
const SERVER_URL = 'https://YOUR-SERVER/api/refresh?token=YOUR_TOKEN';
function onSheetEdit() {
  try { UrlFetchApp.fetch(SERVER_URL, { muteHttpExceptions: true }); } catch (e) {}
}
```
This needs the server to be reachable from the internet (e.g. deployed on Render/Railway/a VPS).
On a local PC, rely on the 60 s sync plus the "Sync now" button.

## Files
- `config.js`    – categories and keyword rules (unchanged from the Apps Script)
- `extractor.js` – extraction logic (ported 1:1)
- `db.js`        – read-only live sheet mirror (in-memory, background sync)
- `server.js`    – Express API: `POST /api/extract`, `GET /api/health`, `/api/refresh`
- `public/index.html` – upload UI with sync status
