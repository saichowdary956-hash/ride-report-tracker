# Deploy On Vercel

This repo includes `vercel.json` and `api/index.py` so Vercel can run the app as a Python serverless function.

## Steps

1. Open Vercel.
2. Import GitHub repo:
   `saichowdary956-hash/ride-report-tracker`
3. Keep the default project settings.
4. Add environment variables:

```text
DATABASE_URL=<your PostgreSQL connection string>
ALLOW_SQLITE_FALLBACK=0
```

5. Deploy.

## Important

Vercel's filesystem is temporary. The app uses `/tmp` only for generated Excel downloads and uploaded batch files during a request. Real tracker data must be stored in PostgreSQL through `DATABASE_URL`.

For lakhs of CSV files, use PostgreSQL, not fallback SQLite. Keep:

```text
ALLOW_SQLITE_FALLBACK=0
```

The app stores indexed `vehicle`, `source_file`, and `drive_id` columns in PostgreSQL so vehicle filtering and CSV-file lists remain practical as data grows.

Vercel cannot open Microsoft Excel on your computer. Use **Edit Excel File** in the app for browser editing, then use **Download Excel Copy** when you need an `.xlsx`.
