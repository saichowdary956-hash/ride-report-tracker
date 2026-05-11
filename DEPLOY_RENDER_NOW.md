# Deploy On Render

Use this link:

https://render.com/deploy?repo=https://github.com/saichowdary956-hash/ride-report-tracker

Render will read `render.yaml` and create:

- `ride-report-tracker` web app
- `ride-report-db` PostgreSQL database
- `DATABASE_URL` connection between them

## Clicks

1. Open the link above.
2. Sign in to Render.
3. Confirm the Blueprint.
4. Click **Apply** or **Deploy Blueprint**.
5. Wait until both the database and web service say live.
6. Open the web service URL.

## Important

The Blueprint uses Render's free database plan so the app works immediately. Free Render Postgres does not include backups. For real long-term/high-volume tracker data, upgrade the database plan before loading important production data.

## If The App Opens With A Storage Warning

That means the web app is live but PostgreSQL is not connected yet. In Render, open the Blueprint/Environment page and sync the Blueprint again so `DATABASE_URL` comes from `ride-report-db`.
