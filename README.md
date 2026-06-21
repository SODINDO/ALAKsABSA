# ABSA Retail Sales Dashboard

A Vite-powered, single-page dashboard built with vanilla JS and Chart.js.
Reads data from an Excel workbook bundled inside the project itself —
no external network call, no CORS issues, no API keys, no Google account
permissions required.

## Project structure

```
absa-vite/
├── data.xlsx          ← the spreadsheet, lives in the project root
├── index.html
├── package.json
├── vite.config.js
├── vercel.json
└── src/
    └── main.js          imports data.xlsx directly at build time
```

## Local development

```bash
npm install
npm run dev
```

Open the URL Vite prints (usually http://localhost:5173).

## Build

```bash
npm run build
```

Output goes to `dist/`. Verify locally with:

```bash
npm run preview
```

## Deploying to Vercel

This repo includes `vercel.json` already configured:

```json
{
  "buildCommand": "npm run build",
  "outputDirectory": "dist",
  "installCommand": "npm install",
  "framework": "vite"
}
```

Import the repo (or drag-and-drop the folder) at vercel.com/new — no
manual configuration needed.

## Updating the data

The Excel file lives at `data.xlsx` in the project root, right next to
the `src/` folder. It is imported directly into the JavaScript bundle
at build time:

```javascript
import dataXlsxUrl from '../data.xlsx?url';
```

Vite resolves this import, copies the file into the build output, and
gives the app a working URL to fetch it from — so there's no need to
keep it inside `public/`.

The dashboard reads it at page load and again every 60 seconds. To
update the numbers:

1. Open `data.xlsx` in the project root using Excel, Google Sheets, or
   LibreOffice.
2. Replace the rows with your real export — keep the same header row.
3. Save the file back to exactly `data.xlsx` in the project root.
4. Run `npm run build` and redeploy (push to your connected Git repo, or
   re-upload to Vercel).

There is no live external link involved — the spreadsheet travels with
the deployed app as a bundled asset, so it always loads instantly and
never depends on network access to Google, a proxy, or any third party.

### Expected columns

The column auto-detector accepts several common header spellings per
field:

| Purpose          | Accepted header names                                   |
|-------------------|------------------------------------------------------------|
| Region/Cluster    | REGION, Region, region, CLUSTER                              |
| Premium amount    | PREMIUM, Premium, premium, PREMIUM_COLLECTED, premium_collected |
| Product           | PRODUCT, Product, product                                    |
| Month             | Month, MONTH, month (values: Jan, Feb, March, April, etc.)   |
| Premium term      | Premiumpayingterm, TERM, premium_paying_term, PPT             |
| Branch            | SALESBRANCH, salesbranch, Branch, BRANCH, branch              |

The starter `data.xlsx` shipped in this project uses exactly these
headers — REGION, SALESBRANCH, PRODUCT, PREMIUM, Month,
Premiumpayingterm — so the simplest path is to keep that header row
and just replace the rows beneath it with your real transactional data.

REGION values must match one of the existing cluster names already
configured in the dashboard (Coast, Nairobi West, Rift Valley, Queensway,
Alternative Channels, Nairobi Central, Premier, Nairobi East, Western,
Mt. Kenya, TMU) for cluster-level rollups to populate; unmatched regions
are simply skipped for that view but still count toward branch-level and
product-level totals.

If your headers differ from all of the above, either rename them in the
spreadsheet or edit the find(...) calls inside rowsToDB() in
src/main.js to add your variant.
