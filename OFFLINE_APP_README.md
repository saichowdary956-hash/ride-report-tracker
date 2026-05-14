# RideReport Offline App

Use `Start_RideReport_Offline.bat` to run the tracker without Render, Vercel, Neon, or internet.

## What Gets Saved

All app data is saved automatically inside:

```text
app_data\outputs\
```

Important files:

- `daily_tracker_backup.sqlite` stores vehicles, uploaded CSV files, tracker rows, settings, and edits.
- `daily_tracker_<vehicle>.xlsx` is regenerated from saved data.
- Downloaded Excel copies include the current date in the filename.

Uploads, deletes, edits, vehicle changes, planned-hour changes, Progress, and Charts are saved immediately to SQLite. You do not need a separate Save button. After closing the browser or app window, run `Start_RideReport_Offline.bat` again and the data will still be there.

## First Run

The first run may install `openpyxl` if it is not already installed. After that, the app works offline.

## Backup

To back up everything, copy the whole `app_data` folder somewhere safe. Restoring is just copying that folder back next to `Start_RideReport_Offline.bat`.
