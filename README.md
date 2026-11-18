# CFS Order Intelligence Dashboard

Responsive React + Vite dashboard for Project and Service Delivery Managers. It loads the included baseline Excel workbooks in the browser, accepts daily Rework Handover uploads, merges records, stores the current view in IndexedDB, and refreshes KPIs immediately.

## Run locally

```bash
npm install
npm run dev
```

## Deploy with GitHub and Vercel

1. Create a GitHub repository and push this folder.
2. In Vercel, select **Add New > Project** and import the repository.
3. Vercel detects Vite automatically. Use build command `npm run build` and output directory `dist`.
4. Deploy. No server or environment variable is required.

## Data behavior

- Baseline files are under `public/baseline`.
- Daily `.xlsx` files are parsed locally in the browser. Workbook data does not leave the device.
- Uploaded rows are normalized to a common model, then merged by source + SO/DDL/serial fallback key.
- Browser state is stored in IndexedDB, which is safer for larger datasets than localStorage.
- Use **Reset local data** to restore the bundled baseline.

## AI Order Assistant

The assistant is an offline natural-language query layer over the records in memory. It supports order, DDL, serial, customer, status, aging, region and summary questions. No workbook content is sent to an external AI service.

Examples:
- `Show order G001081500`
- `How many pending repair items?`
- `Summarise EPA rework`
- `Show items older than 30 days`
- `Top customers`
