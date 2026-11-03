# Faboolux Materials Extractor (no server)

A static web page. Drop an estimate Excel and the materials table appears instantly.
Everything runs **inside the browser**: there is no server, nothing is uploaded or saved,
and nothing is ever written to Google Sheets.

## How the data stays in sync
- The page **reads** your database Google Sheet (`Cabinet_Data`, `Laminate_Data`, `masterdata`) with a read-only Google key.
- It re-checks the sheet **every 60 seconds**, whenever you switch back to the tab, and when you press **Sync now**.
- If something changed in the sheet, the data updates and **the estimate you already uploaded is re-processed automatically**.
- The last copy is saved in your browser, so the page starts instantly and still works if Google is briefly unreachable.

## One-time setup (about 5 minutes)

### 1. Share the database sheet (read-only)
Google Sheet -> **Share** -> General access -> **Anyone with the link** -> **Viewer**.

### 2. Create a read-only Google API key
1. Open https://console.cloud.google.com/ and create a project (any name).
2. **APIs & Services -> Library -> "Google Sheets API" -> Enable.**
3. **APIs & Services -> Credentials -> Create credentials -> API key.**
4. Click the key -> **Restrict key**:
   - *Websites (HTTP referrers)*: add `https://fabooluxinterior.github.io/*`
     (and `http://localhost:*` if you want to test locally)
   - *API restrictions*: **Restrict key -> Google Sheets API**
5. Copy the key and paste it into `settings.js` as `API_KEY`, then commit.

The key can only *read* sheets that are shared with a link, and only when called from your website,
so it is safe to keep in the repository.

### 3. Publish the page on GitHub Pages (free)
Repository -> **Settings -> Pages -> Build and deployment -> Source: Deploy from a branch ->
Branch: `main`, folder `/ (root)` -> Save.**
After a minute the app is live at https://fabooluxinterior.github.io/Material_extractor/

## Files
- `index.html`   - the page
- `app.js`       - upload, rendering, auto re-processing
- `sheet.js`     - read-only live sheet sync (60 s, on focus, on demand)
- `extractor.js` - extraction rules, ported 1:1 from the Apps Script
- `config.js`    - categories and keyword rules
- `settings.js`  - sheet ID, API key, sync interval
