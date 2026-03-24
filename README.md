# DAILA Word Add-in

Three-button task pane for Word — Document Summary, Contract Analysis, Issue List Generation.
Each action reads the open document, posts it to a Make webhook, and opens the result as a new Word document.

---

## Files

```
daila-addin/
  taskpane.html   ← the task pane UI
  taskpane.js     ← all logic (edit this first)
  manifest.xml    ← Office add-in manifest
  assets/
    icon-16.png   ← you need to add these
    icon-32.png
    icon-80.png
```

---

## Step 1 — Add your Make webhook URLs

Open `taskpane.js` and replace the three placeholder URLs at the top:

```js
const WEBHOOKS = {
  summary:  'https://hook.eu2.make.com/YOUR_SUMMARY_HOOK',
  contract: 'https://hook.eu2.make.com/YOUR_CONTRACT_HOOK',
  issues:   'https://hook.eu2.make.com/YOUR_ISSUES_HOOK',
};
```

---

## Step 2 — What Make receives

Your Make webhook receives a POST with this JSON body:

```json
{
  "action":    "summary",          // "summary" | "contract" | "issues"
  "filename":  "ClientMatter.docx",
  "document":  "<base64 string>",  // the full .docx file as base64
  "timestamp": "2026-03-24T10:00:00.000Z"
}
```

In Make: use the `document` field as the file input for your document upload module.
Set the filename to the `filename` field value.

---

## Step 3 — What Make must return

Make's webhook response must be JSON in one of these two shapes:

**Shape A — base64 (recommended, opens directly in Word):**
```json
{
  "type":     "base64",
  "data":     "<base64 encoded .docx>",
  "filename": "ClientMatter_Summary.docx",
  "message":  "Summary generated successfully."
}
```

**Shape B — URL (shows a download button instead):**
```json
{
  "type":     "url",
  "url":      "https://your-storage.com/outputs/summary.docx",
  "filename": "ClientMatter_Summary.docx",
  "message":  "Summary generated successfully."
}
```

> Shape A opens the result directly as a new Word document.
> Shape B shows a download button in the task pane.
> Shape A is strongly recommended for the demo.

---

## Step 4 — Host the add-in files

The two files (`taskpane.html` and `taskpane.js`) must be served over **HTTPS**.

**Quickest option for testing: GitHub Pages**
1. Create a new GitHub repo (public or private)
2. Add `taskpane.html`, `taskpane.js`, and an `assets/` folder with placeholder icons
3. Enable GitHub Pages (Settings → Pages → Deploy from main branch)
4. Your URL will be `https://YOUR-USERNAME.github.io/YOUR-REPO/taskpane.html`

**Other options:** Vercel, Netlify, Cloudflare Pages — all free and serve HTTPS automatically.

Once hosted, replace all instances of `https://YOUR-HOST` in `manifest.xml` with your actual URL.

---

## Step 5 — Sideload for testing (no App Store needed)

### Word Desktop (Windows)
1. Open Word → File → Options → Trust Center → Trust Center Settings
2. Trusted Add-in Catalogs → add a **network share** path pointing to your `manifest.xml` folder
   (e.g. `\\localhost\daila-addin` or any shared folder)
3. Restart Word → Insert → My Add-ins → Shared Folder → DAILA

### Word Desktop (Mac)
1. Copy `manifest.xml` to:
   `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/`
2. Restart Word → Insert → My Add-ins → Developer Add-ins → DAILA

### Word Online (any browser)
1. Open any Word doc in Word Online
2. Insert → Add-ins → Upload My Add-in → browse to `manifest.xml`

---

## Step 6 — Deploy to a law firm (optional)

For a firm-wide rollout (no sideloading):
1. Host the files on HTTPS (GitHub Pages, Vercel, etc.)
2. Give the firm's Microsoft 365 admin your `manifest.xml`
3. Admin goes to: Microsoft 365 Admin Centre → Settings → Integrated Apps → Upload custom app
4. The DAILA button appears in Word's Home ribbon for all users — no installation needed

Requires Microsoft 365 Business Basic or above (not Family).

---

## Make scenario setup notes

Each scenario needs to:
1. Receive the webhook (Custom webhook trigger module)
2. Decode the base64 `document` field into a file
3. Upload to DAILA (`/api/v1/documents/upload` or equivalent)
4. Call the appropriate DAILA endpoint (summarise / contract analysis / issue list)
5. Retrieve the output `.docx`
6. Return it as base64 in the webhook response

For step 5 → 6, use Make's "Webhook response" module with:
- Content type: `application/json`
- Body: `{"type":"base64","data":"{{base64OutputDoc}}","filename":"{{filename}}_Summary.docx"}`

Make's timeout for synchronous webhook responses is **40 seconds**.
If DAILA takes longer, use Shape B (URL): store the output file, return the URL,
and let the task pane offer a download button.
