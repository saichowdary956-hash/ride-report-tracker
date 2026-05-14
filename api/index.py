import os
import sys
import traceback
import csv
import io
from pathlib import Path

os.environ.setdefault("RIDE_REPORT_BASE_DIR", "/tmp/ride-report-tracker")

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from flask import Flask, redirect, request, send_file

from ride_report_app import (
    DEFAULT_DOWNLOADS,
    OUTPUT_DIR,
    UPLOAD_DIR,
    DEFAULT_TOTAL_PLANNED_HOURS,
    normalize_vehicle,
    page,
    row_from_form,
    tracker_path_for_vehicle,
    vehicle_setting_key,
)
from ride_report_tool import (
    add_database_row,
    add_vehicle,
    delete_database_row,
    delete_database_rows_by_source_files,
    import_tracker_workbook,
    load_uploaded_csv_file,
    reconstructed_csv_from_rows,
    process_reports,
    refresh_uploaded_csv_from_rows,
    rebuild_tracker_from_database,
    remove_vehicle,
    save_uploaded_csv_file,
    set_setting,
    update_database_row,
    uploaded_csv_files_from_database,
)

app = Flask(__name__)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)


def setup_error_page(exc):
    error_text = traceback.format_exc(limit=4)
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>RideReport Setup Error</title>
  <style>
    body {{ font-family: Segoe UI, Arial, sans-serif; background: #f4f7f9; color: #17212b; margin: 0; padding: 32px; }}
    main {{ max-width: 900px; margin: 0 auto; background: #fff; border: 1px solid #d8dee6; border-radius: 8px; padding: 22px; }}
    code, pre {{ background: #eef3f6; border-radius: 6px; padding: 3px 5px; }}
    pre {{ overflow: auto; padding: 12px; white-space: pre-wrap; }}
  </style>
</head>
<body>
  <main>
    <h1>Database setup needs attention</h1>
    <p>The app is deployed, but it could not connect to PostgreSQL.</p>
    <p>In Vercel, connect Neon Postgres under <strong>Storage</strong>. The app now accepts any of these variables:</p>
    <ul>
      <li><code>DATABASE_URL</code></li>
      <li><code>POSTGRES_URL</code></li>
      <li><code>POSTGRES_PRISMA_URL</code></li>
      <li><code>POSTGRES_URL_NON_POOLING</code></li>
    </ul>
    <p>Keep <code>ALLOW_SQLITE_FALLBACK=0</code> only after the Neon variable exists.</p>
    <h2>Error</h2>
    <pre>{str(exc)}</pre>
    <h2>Trace</h2>
    <pre>{error_text}</pre>
  </main>
</body>
</html>""", 500


@app.errorhandler(Exception)
def handle_error(exc):
    return setup_error_page(exc)


def form_lists():
    return {key: request.form.getlist(key) for key in request.form.keys()}


def request_vehicle(default="Default"):
    return normalize_vehicle(request.values.get("vehicle") or default)


def render_app(message="", processed=None, skipped=None, pending_folder="", active_tab="home", selected_source="", selected_csv_action="", vehicle=None):
    return page(
        message,
        processed or [],
        skipped or [],
        pending_folder=pending_folder,
        active_tab=active_tab,
        selected_source=selected_source,
        selected_csv_action=selected_csv_action,
        vehicle=vehicle or request_vehicle(),
    )


@app.get("/")
def home():
    vehicle = request_vehicle()
    return render_app(
        active_tab=request.args.get("tab", "home"),
        selected_source=request.args.get("source", ""),
        selected_csv_action=request.args.get("csv_action", ""),
        vehicle=vehicle,
    )


@app.get("/download")
def download():
    vehicle = request_vehicle()
    target = tracker_path_for_vehicle(vehicle)
    rebuild_tracker_from_database(OUTPUT_DIR, tracker_name=target.name, vehicle=vehicle)
    return send_file(target, as_attachment=True, download_name=target.name)


@app.get("/download-uploaded-csv-list")
def download_uploaded_csv_list():
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Vehicle", "CSV File", "Uploaded At", "Updated At"])
    vehicle = request_vehicle()
    for item in uploaded_csv_files_from_database(OUTPUT_DIR, vehicle=vehicle):
        writer.writerow([item.get("vehicle", ""), item.get("source_file", ""), item.get("uploaded_at", ""), item.get("updated_at", "")])
    data = output.getvalue().encode("utf-8-sig")
    return app.response_class(
        data,
        mimetype="text/csv",
        headers={"Content-Disposition": f'attachment; filename="uploaded_csv_files_{vehicle}.csv"'},
    )


@app.get("/download-uploaded-csv")
def download_uploaded_csv():
    source = request.args.get("source", "")
    vehicle = request_vehicle()
    stored = load_uploaded_csv_file(OUTPUT_DIR, vehicle, source)
    if not stored:
        stored = reconstructed_csv_from_rows(OUTPUT_DIR, vehicle, source)
    if not stored:
        return "Uploaded CSV not found", 404
    filename, data = stored
    return app.response_class(
        data,
        mimetype="text/csv",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.get("/open-excel")
def open_excel():
    return redirect("/?tab=excel-editor")


@app.post("/upload")
def upload():
    vehicle = request_vehicle()
    staging = UPLOAD_DIR / "_last_batch"
    staging.mkdir(parents=True, exist_ok=True)
    for old_file in staging.glob("*.csv"):
        old_file.unlink()
    csv_paths = []
    for uploaded in request.files.getlist("files"):
        if uploaded and uploaded.filename.lower().endswith(".csv"):
            target = staging / Path(uploaded.filename).name
            uploaded.save(target)
            save_uploaded_csv_file(OUTPUT_DIR, vehicle, target.name, target.read_bytes())
            csv_paths.append(target)
    if not csv_paths:
        return render_app("No CSV files were received. Please choose one or more .csv files.", active_tab="home", vehicle=vehicle)
    result = process_reports(staging, OUTPUT_DIR, tracker_name=tracker_path_for_vehicle(vehicle).name, vehicle=vehicle)
    return render_app(result["message"], result["processed"], result["skipped"], active_tab="home", vehicle=vehicle)


@app.post("/upload-anyway")
def upload_anyway():
    vehicle = request_vehicle()
    staging = UPLOAD_DIR / "_last_batch"
    result = process_reports(staging, OUTPUT_DIR, tracker_name=tracker_path_for_vehicle(vehicle).name, allow_partial=True, vehicle=vehicle)
    return render_app("Processed with missing fields left blank. " + result["message"], result["processed"], result["skipped"], active_tab="home", vehicle=vehicle)


@app.post("/folder")
def folder():
    vehicle = request_vehicle()
    form = form_lists()
    folder_path = Path(form.get("folder", [str(DEFAULT_DOWNLOADS)])[0])
    allow_partial = form.get("allow_partial", ["0"])[0] == "1"
    result = process_reports(folder_path, OUTPUT_DIR, tracker_name=tracker_path_for_vehicle(vehicle).name, allow_partial=allow_partial, vehicle=vehicle)
    prefix = "Processed with missing fields left blank. " if allow_partial else ""
    return render_app(prefix + result["message"] + f" Source: {folder_path}.", result["processed"], result["skipped"], pending_folder=str(folder_path), active_tab="home", vehicle=vehicle)


@app.post("/row/update")
def row_update():
    vehicle = request_vehicle()
    form = form_lists()
    row = row_from_form(form)
    row["Vehicle"] = vehicle
    update_database_row(OUTPUT_DIR, form.get("row_id", [""])[0], row)
    refresh_uploaded_csv_from_rows(OUTPUT_DIR, vehicle, row.get("Source File", ""))
    rebuild_tracker_from_database(OUTPUT_DIR, tracker_name=tracker_path_for_vehicle(vehicle).name, vehicle=vehicle)
    active_tab = form.get("return_tab", ["excel-editor"])[0] or "excel-editor"
    selected_source = form.get("source", [""])[0]
    selected_action = "edit" if active_tab == "csv-list" and selected_source else ""
    return render_app(
        "Row updated in the stored CSV data and the Excel tracker.",
        active_tab=active_tab,
        selected_source=selected_source,
        selected_csv_action=selected_action,
        vehicle=vehicle,
    )


@app.post("/row/delete")
def row_delete():
    vehicle = request_vehicle()
    form = form_lists()
    delete_database_row(OUTPUT_DIR, form.get("row_id", [""])[0])
    rebuild_tracker_from_database(OUTPUT_DIR, tracker_name=tracker_path_for_vehicle(vehicle).name, vehicle=vehicle)
    return render_app("Row deleted from the tracker database.", active_tab="excel-editor", vehicle=vehicle)


@app.post("/row/add")
def row_add():
    vehicle = request_vehicle()
    row = row_from_form(form_lists())
    row["Vehicle"] = vehicle
    add_database_row(OUTPUT_DIR, row)
    rebuild_tracker_from_database(OUTPUT_DIR, tracker_name=tracker_path_for_vehicle(vehicle).name, vehicle=vehicle)
    return render_app("Row added to the tracker database.", active_tab="excel-editor", vehicle=vehicle)


@app.post("/sync-excel")
def sync_excel():
    vehicle = request_vehicle()
    rows = import_tracker_workbook(OUTPUT_DIR, tracker_path_for_vehicle(vehicle), vehicle=vehicle)
    return render_app(f"Synced {len(rows)} Excel row(s) to the tracker database.", active_tab="excel-editor", vehicle=vehicle)


@app.post("/refresh-totals")
def refresh_totals():
    vehicle = request_vehicle()
    return render_app("Refresh Totals was removed because totals update automatically when you upload, edit, or delete CSV data.", active_tab="excel-editor", vehicle=vehicle)


@app.post("/delete-csv")
def delete_csv():
    vehicle = request_vehicle()
    selected_sources = request.form.getlist("source_file")
    deleted = delete_database_rows_by_source_files(OUTPUT_DIR, selected_sources, vehicle=vehicle)
    rebuild_tracker_from_database(OUTPUT_DIR, tracker_name=tracker_path_for_vehicle(vehicle).name, vehicle=vehicle)
    message = f"Deleted {deleted} Daily Tracker row(s) from selected CSV file(s)." if selected_sources else "No CSV files selected for deletion."
    return render_app(message, active_tab="home", vehicle=vehicle)


@app.post("/settings/planned-hours")
def planned_hours():
    vehicle = request_vehicle()
    raw_value = request.form.get("total_planned_hours", str(DEFAULT_TOTAL_PLANNED_HOURS)).strip()
    try:
        planned = float(raw_value)
        if planned < 0:
            raise ValueError
        set_setting(OUTPUT_DIR, vehicle_setting_key(vehicle, "total_planned_hours"), planned)
        message = f"Total planned hours updated to {planned:g}."
    except ValueError:
        message = "Please enter a valid non-negative planned hours value."
    return render_app(message, active_tab="progress", vehicle=vehicle)


@app.post("/vehicle/select")
def vehicle_select():
    vehicle = request_vehicle()
    rebuild_tracker_from_database(OUTPUT_DIR, tracker_name=tracker_path_for_vehicle(vehicle).name, vehicle=vehicle)
    return render_app(f"Switched to vehicle: {vehicle}.", active_tab="home", vehicle=vehicle)


@app.post("/vehicle/add")
def vehicle_add():
    vehicle = request.form.get("vehicle", "").strip()
    if vehicle:
        add_vehicle(OUTPUT_DIR, vehicle)
        set_setting(OUTPUT_DIR, vehicle_setting_key(vehicle, "total_planned_hours"), DEFAULT_TOTAL_PLANNED_HOURS)
        rebuild_tracker_from_database(OUTPUT_DIR, tracker_name=tracker_path_for_vehicle(vehicle).name, vehicle=vehicle)
        message = f"Added and switched to vehicle: {vehicle}."
    else:
        message = "Please enter a vehicle name."
    return render_app(message, active_tab="home", vehicle=vehicle or "Default")


@app.post("/vehicle/remove")
def vehicle_remove():
    vehicle = request.form.get("vehicle", active_vehicle()).strip()
    if vehicle and vehicle != "Default":
        deleted = remove_vehicle(OUTPUT_DIR, vehicle)
        rebuild_tracker_from_database(OUTPUT_DIR, tracker_name=tracker_path_for_vehicle("Default").name, vehicle="Default")
        message = f"Removed vehicle '{vehicle}' and deleted {deleted} row(s). Switched back to Default."
    else:
        message = "Default vehicle cannot be removed."
    return render_app(message, active_tab="home", vehicle="Default" if vehicle != "Default" else vehicle)
