# CFS Operations Intelligence
React + Vite management dashboard ready for GitHub -> Vercel deployment.

## Features
- Professional dark CFS operations command-center layout.
- Dedicated Operations Overview, Rework Operations, Order Aging & Fulfilment, DOA Management, Actions & Commitments tabs.
- Daily Rework Handover upload: the latest successful rework upload replaces only the Rework dataset; other workbooks stay intact.
- Previous Rework snapshot retained in IndexedDB before replacement.
- All workbook data persists in this browser using IndexedDB until replaced.
- Charts, management KPIs, global table search, and minimizable AI assistant.

## Deploy update
Extract this package over your existing local repository, then run:

```powershell
git add .
git commit -m "Upgrade CFS Operations Intelligence dashboard"
git push origin main
```

Vercel should redeploy automatically from GitHub.

## Important production note
IndexedDB is device/browser-local. For a multi-user Lenovo deployment, move the shared dataset to a Lenovo-approved authenticated data store so all PM/SDM users see the same current snapshot.
