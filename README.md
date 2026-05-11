# RideReport Vehicle Tracker

This app uploads RideReport CSV files, tracks each vehicle separately, calculates completed hours by condition, and rebuilds a Daily Tracker Excel file from the stored data.

## Local Run

```powershell
pip install -r requirements.txt
python .\ride_report_app.py --open
```

Then open:

```text
http://127.0.0.1:8765
```

Local mode uses:

- `outputs\daily_tracker_backup.sqlite` for backup storage
- `outputs\daily_tracker_<Vehicle>.xlsx` for Excel exports
- `uploads\` for temporary uploaded CSV files

## Cloud Storage

Set `DATABASE_URL` to a PostgreSQL connection string to use cloud storage instead of the local SQLite file.

```powershell
$env:DATABASE_URL="postgresql://USER:PASSWORD@HOST:5432/postgres"
python .\ride_report_app.py
```

When `DATABASE_URL` is present, all vehicle rows, uploaded CSV records, settings, and progress data are stored in PostgreSQL. The app has no built-in CSV count limit per vehicle; storage is limited by your database plan.

## Render Deployment

1. Push this folder to a GitHub repository.
2. Create a PostgreSQL database, for example Supabase or Render Postgres.
3. Copy the database connection string.
4. In Render, create a new Web Service from the GitHub repository.
5. Use:
   - Build command: `pip install -r requirements.txt`
   - Start command: `python ride_report_app.py`
6. Add environment variable:
   - `DATABASE_URL=<your postgres connection string>`

`render.yaml` is included so Render can detect the Python service settings.

## Notes

- Select the active vehicle before uploading CSV files.
- The upload button shows the target vehicle, for example `Upload CSV Files to Jeep`.
- Excel files are generated from database data. In cloud deployment, download the Excel file instead of trying to open Excel on the server.
- Keep `.env`, `outputs/`, and `uploads/` out of GitHub. They are ignored by `.gitignore`.
