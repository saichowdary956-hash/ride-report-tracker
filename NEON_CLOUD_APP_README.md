# RideReport Neon Cloud App

Use `Start_RideReport_Neon_Cloud.bat` when you want the desktop app to store data in Neon Postgres.

## Important

Do not use `Start_RideReport_Offline.bat` when you want Neon storage. The offline launcher intentionally uses local SQLite only.

## First Run

1. Double-click `Start_RideReport_Neon_Cloud.bat`.
2. Paste your Neon connection string when asked.
3. The launcher saves it in `neon_database_url.txt` next to the app.
4. The app opens in your browser.

The launcher sets `ALLOW_SQLITE_FALLBACK=0`, so if Neon is not reachable the app will show an error instead of silently saving to local SQLite.

## File Storage

Inside the app, open the `File Storage` tab for the selected vehicle.

The tab shows database-backed folders:

- `/<vehicle>/csv` for uploaded CSV files
- `/<vehicle>/excel` for the current generated Excel workbook

These files are stored as bytes in Neon Postgres. Uploading CSV files, editing rows, deleting files, and rebuilding Excel all update Neon immediately.

## Backup

Your main backup is Neon. You can also download CSV and Excel files from the app's `File Storage` tab.
