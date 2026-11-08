# Inventory Console

A browser-only admin dashboard for CSV, Excel and Google Sheets data. It has no backend and no build step.

## Run
```bash
./start.sh            # serves on http://localhost:8080 and opens the browser
```
You can also open `index.html` directly.

## Data sources
- **Upload**: drag a `.csv`, `.tsv`, `.xlsx` or `.xls` file anywhere onto the page. Excel quirks such as the `sep=;` line and `="0123"` text cells are handled.
- **Google Sheets**: in Google Sheets, choose *Share → Anyone with the link (Viewer)*, or *File → Share → Publish to web*. Paste the link into **Data sources**. Auto-refresh can be set from 1 to 60 minutes.
- **Merging uploads (default)**: each new file is merged into the saved data, unique by **Booking Number**. If a booking is already saved, all of its rows are replaced by its rows in the newest file. Bookings that are not in the new file are kept. Upload 400 rows today and 500 overlapping rows tomorrow, and you get the 500 unique entries with the latest data. Columns are matched by name, so their order can differ. *Data sources → Upload history* lists what each file added or updated. Switch *New uploads* to **Replace** if you want the old behaviour.
- The merged data is saved in this browser (IndexedDB), so it survives reloads and restarts. It stays until you click **Remove data**. Loading a Google Sheet asks before replacing saved uploads. Sheet links, hidden columns and theme are kept in localStorage.

## Features
- **Overview**: KPIs and charts. Click any bar or slice to filter on it.
- **Data table**: full-text search, sortable columns, column picker, paging, row detail drawer.
- **Filters panel**: checkbox facets with live counts, plus date and number ranges with quick presets.
- **Search syntax**: `jakarta`, `pod:jakarta`, `status:cancel`, `-unassigned`, `"port kelang"`.
- **Pivot**: rows × columns × (count / unique count / sum / average). Click a cell to drill down.
- **Export**: the filtered rows go to CSV or Excel.

### Shipping bookings reports
When the file has *Booking Number* and *Container Number* columns, more features turn on:
- **Bookings view**: one row per booking, with container counts, readiness and the latest event.
- **Voyages view**: per vessel and voyage, the ETD and how many containers have reached each stage (empty picked up, gated in, at POL, sailed).
- **Derived columns**: ETD Type (Actual/Estimated), Container Assigned, TEU, Milestone.
- **Booking drawer**: a milestone timeline for each container.

## Files
- `index.html`, `css/styles.css`, `js/app.js`: the app
- `vendor/`: PapaParse, SheetJS and Chart.js, stored locally so the app works offline
