from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote_plus, urlparse
import argparse
import csv
from datetime import date
import html
import io
import mimetypes
import os
import webbrowser

from ride_report_tool import (
    BURNDOWN_ROWS,
    DAILY_TRACKER_COLUMNS,
    add_vehicle,
    add_database_row,
    count_rows_in_database,
    delete_database_row,
    delete_database_rows_by_source_files,
    database_fallback_reason,
    import_tracker_workbook,
    load_rows_from_database,
    process_reports,
    rebuild_tracker_from_database,
    get_setting,
    remove_vehicle,
    load_uploaded_csv_file,
    refresh_uploaded_csv_from_rows,
    reconstructed_csv_from_rows,
    save_uploaded_csv_file,
    set_setting,
    source_files_from_database,
    sync_database_from_tracker,
    totals_from_daily_rows,
    update_database_row,
    uploaded_csv_files_from_database,
    vehicles_from_database,
    safe_load_workbook,
)


DEFAULT_TOTAL_PLANNED_HOURS = 300
PROGRESS_FIELD_BY_CONDITION = {
    "Sunny": "Sunny",
    "Low Sun": "Low Sun",
    "Cloudy": "Cloudy",
    "Rain": "Rain",
    "Fog": "Fog",
    "Snow": "Snow",
    "City (intense traffic)": "City (intense traffic)",
    "Country": "Country",
    "Highway": "Highway",
    "Construction Site": "Construction",
    "Tunnel*": "Tunnel",
    "Day": "Day",
    "Dawn": "Dawn",
    "Lit Night": "Lit Night",
    "Dark Night": "Dark Night",
    "Flow": "Flow",
    "Jam": "Jam",
    "5-30 km/h (3-18 mph)": "3-18 mph",
    "30-60 km/h (18-37 mph)": "19-37 mph",
    "60-90 km/h (37-55 mph)": "38-55 mph",
    "90-130 km/h (55-80 mph)": "56-80 mph",
    "130-250 km/h (80-155 mph)*": "81-155 mph",
}


BASE_DIR = Path(os.environ.get("RIDE_REPORT_BASE_DIR", Path(__file__).resolve().parent))
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
TRACKER_PATH = OUTPUT_DIR / "daily_tracker.xlsx"
DEFAULT_DOWNLOADS = Path.home() / "Downloads"


def can_open_excel_locally():
    return hasattr(os, "startfile")


def excel_action_label():
    return "Edit Daily Tracker Excel" if can_open_excel_locally() else "Download Daily Tracker Excel"


def read_preview(path, sheet_name, limit=8):
    if not path.exists():
        return []
    workbook = safe_load_workbook(path, data_only=True)
    if workbook is None:
        return []
    if sheet_name not in workbook.sheetnames:
        return []
    sheet = workbook[sheet_name]
    rows = []
    for values in sheet.iter_rows(max_row=min(sheet.max_row, limit + 1), values_only=True):
        rows.append(["" if value is None else str(value) for value in values])
    return rows


def active_vehicle():
    return get_setting(OUTPUT_DIR, "active_vehicle", "Default") or "Default"


def normalize_vehicle(vehicle):
    return str(vehicle or "").strip() or "Default"


def vehicle_setting_key(vehicle, key):
    return f"vehicle::{vehicle}::{key}"


def vehicle_param(vehicle=None):
    return quote_plus(normalize_vehicle(vehicle or active_vehicle()))


def current_rows(vehicle=None):
    return load_rows_from_database(OUTPUT_DIR, vehicle=normalize_vehicle(vehicle or active_vehicle()))


def tracker_path_for_vehicle(vehicle=None):
    vehicle = normalize_vehicle(vehicle or active_vehicle())
    safe_vehicle = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in vehicle)
    return OUTPUT_DIR / f"daily_tracker_{safe_vehicle}.xlsx"


def dated_excel_filename(path):
    path = Path(path)
    return f"{path.stem}_{date.today().strftime('%Y-%m-%d')}{path.suffix}"


def active_tracker_path():
    return tracker_path_for_vehicle(active_vehicle())


def process_csv_paths(csv_paths, allow_partial=False, vehicle=None):
    vehicle = normalize_vehicle(vehicle or active_vehicle())
    staging = UPLOAD_DIR / "_last_batch"
    if staging.exists():
        for old_file in staging.glob("*.csv"):
            old_file.unlink()
    staging.mkdir(parents=True, exist_ok=True)
    for csv_path in csv_paths:
        target = staging / csv_path.name
        if csv_path.resolve() != target.resolve():
            target.write_bytes(csv_path.read_bytes())
        save_uploaded_csv_file(OUTPUT_DIR, vehicle, target.name, target.read_bytes())
    return process_reports(staging, OUTPUT_DIR, tracker_name=tracker_path_for_vehicle(vehicle).name, allow_partial=allow_partial, vehicle=vehicle)


def uploaded_csv_list_bytes(vehicle):
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Vehicle", "CSV File", "Uploaded At", "Updated At"])
    for item in uploaded_csv_files_from_database(OUTPUT_DIR, vehicle=vehicle):
        writer.writerow([item.get("vehicle", ""), item.get("source_file", ""), item.get("uploaded_at", ""), item.get("updated_at", "")])
    return output.getvalue().encode("utf-8-sig")


def parse_multipart(body, content_type):
    marker = "boundary="
    if marker not in content_type:
        return []
    boundary = ("--" + content_type.split(marker, 1)[1]).encode()
    files = []
    for part in body.split(boundary):
        if b"\r\n\r\n" not in part:
            continue
        header_blob, content = part.split(b"\r\n\r\n", 1)
        headers = header_blob.decode("utf-8", "ignore")
        if 'name="files"' not in headers or "filename=" not in headers:
            continue
        filename = headers.split("filename=", 1)[1].split("\r\n", 1)[0].strip('"')
        content = content.rstrip(b"\r\n-")
        if filename.lower().endswith(".csv") and content:
            UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
            target = UPLOAD_DIR / Path(filename).name
            target.write_bytes(content)
            files.append(target)
    return files


def tracker_table_html(rows, vehicle=None, return_tab="excel-editor", source=""):
    if not rows:
        return "<p class='empty'>No rows yet.</p>"
    headers = [column[0] for column in DAILY_TRACKER_COLUMNS]
    header = "".join(f"<th>{html.escape(value)}</th>" for value in headers) + "<th>Actions</th>"
    body_rows = []
    for row in rows:
        row_id = html.escape(str(row.get("_id", "")))
        cells = ""
        for header_name in headers:
            value = html.escape(str(row.get(header_name, "")))
            field_class = "field-drive-id" if header_name == "Drive ID" else ""
            cells += (
                "<td>"
                f"<input class=\"{field_class}\" form=\"edit-{row_id}\" name=\"{html.escape(header_name)}\" value=\"{value}\">"
                "</td>"
            )
        actions = f"""
        <td class="actions">
          <form id="edit-{row_id}" action="/row/update?vehicle={vehicle_param(vehicle)}" method="post">
            <input type="hidden" name="row_id" value="{row_id}">
            <input type="hidden" name="return_tab" value="{html.escape(return_tab)}">
            <input type="hidden" name="source" value="{html.escape(source)}">
            <button type="submit">Update</button>
          </form>
          <form action="/row/delete?vehicle={vehicle_param(vehicle)}" method="post">
            <input type="hidden" name="row_id" value="{row_id}">
            <button class="danger" type="submit">Delete</button>
          </form>
        </td>"""
        body_rows.append(f"<tr>{cells}{actions}</tr>")
    return f"<table><thead><tr>{header}</tr></thead><tbody>{''.join(body_rows)}</tbody></table>"


def preview_table_html(rows):
    if not rows:
        return "<p class='empty'>Open or sync the Excel tracker to preview rows here.</p>"
    head, body = rows[0], rows[1:]
    header = "".join(f"<th>{html.escape(str(value))}</th>" for value in head)
    body_rows = []
    for row in body:
        cells = "".join(f"<td>{html.escape(str(value))}</td>" for value in row)
        body_rows.append(f"<tr>{cells}</tr>")
    return f"<table><thead><tr>{header}</tr></thead><tbody>{''.join(body_rows)}</tbody></table>"


def daily_tracker_table_html(rows):
    if not rows:
        return "<p class='empty'>No Daily Tracker rows yet.</p>"
    headers = [column[0] for column in DAILY_TRACKER_COLUMNS]
    header = "".join(f"<th>{html.escape(str(value))}</th>" for value in headers)
    body_rows = []
    for row in rows:
        cells = "".join(f"<td>{html.escape(str(row.get(value, '')))}</td>" for value in headers)
        body_rows.append(f"<tr>{cells}</tr>")
    return f"<table><thead><tr>{header}</tr></thead><tbody>{''.join(body_rows)}</tbody></table>"


def add_row_form_html(vehicle=None):
    headers = [column[0] for column in DAILY_TRACKER_COLUMNS]
    inputs = "".join(
        f"<label>{html.escape(header)}<input name=\"{html.escape(header)}\" placeholder=\"{html.escape(header)}\"></label>"
        for header in headers
    )
    return f"""
    <form action="/row/add?vehicle={vehicle_param(vehicle)}" method="post" class="add-grid">
      {inputs}
      <button type="submit">Add Row</button>
    </form>"""


def browser_editor_html(rows, vehicle=None):
    vehicle = normalize_vehicle(vehicle or active_vehicle())
    return f"""
    <div class="editor-actions">
      <a class="button" href="/download?vehicle={vehicle_param(vehicle)}">Download Excel Copy</a>
    </div>
    <div class="muted">Edit values in the table, then click Update on that row. Changes are saved to the tracker database and the Excel file is regenerated.</div>
    <div class="table-wrap editor-table">{tracker_table_html(rows, vehicle)}</div>
    """


def totals_table_html(rows):
    totals = totals_from_daily_rows(rows)
    if not totals:
        return "<p class='empty'>No totals yet.</p>"
    body_rows = []
    for item in totals:
        body_rows.append(
            "<tr>"
            f"<td>{html.escape(item['Category'])}</td>"
            f"<td>{html.escape(item['Field'])}</td>"
            f"<td>{html.escape(item['Total'])}</td>"
            "</tr>"
        )
    return "<table><thead><tr><th>Category</th><th>Field</th><th>Total Completed</th></tr></thead><tbody>" + "".join(body_rows) + "</tbody></table>"


def parse_duration_to_hours(value):
    parts = str(value or "00:00:00").split(":")
    if len(parts) != 3:
        return 0
    try:
        hours, minutes, seconds = [int(float(part)) for part in parts]
    except ValueError:
        return 0
    return hours + minutes / 60 + seconds / 3600


def current_planned_hours(vehicle=None):
    vehicle = normalize_vehicle(vehicle or active_vehicle())
    raw_value = get_setting(OUTPUT_DIR, vehicle_setting_key(vehicle, "total_planned_hours"), str(DEFAULT_TOTAL_PLANNED_HOURS))
    try:
        return float(raw_value)
    except ValueError:
        return float(DEFAULT_TOTAL_PLANNED_HOURS)


def progress_table_html(rows, total_planned_hours):
    totals_by_field = {item["Field"]: item["Total"] for item in totals_from_daily_rows(rows)}
    body_rows = []
    for category, condition, percent in BURNDOWN_ROWS:
        field = PROGRESS_FIELD_BY_CONDITION.get(condition, condition)
        completed_text = totals_by_field.get(field, "00:00:00")
        completed_hours = parse_duration_to_hours(completed_text)
        planned_hours = round(total_planned_hours * percent / 100, 2)
        remaining_hours = round(max(planned_hours - completed_hours, 0), 2)
        body_rows.append(
            "<tr>"
            f"<td>{html.escape(category)}</td>"
            f"<td>{html.escape(condition)}</td>"
            f"<td>{percent}%</td>"
            f"<td>{planned_hours:.2f}</td>"
            f"<td>{html.escape(completed_text)}</td>"
            f"<td>{remaining_hours:.2f}</td>"
            "</tr>"
        )
    return (
        "<table><thead><tr>"
        "<th>Category</th><th>Condition</th><th>%</th><th>Planned Hours</th><th>Completed Hours</th><th>Remaining Pending Hours</th>"
        "</tr></thead><tbody>"
        + "".join(body_rows)
        + "</tbody></table>"
    )


def charts_html(rows, total_planned_hours, vehicle=None):
    total_completed = sum(parse_duration_to_hours(row.get("Overall Session Time", "")) for row in rows)
    total_remaining = max(total_planned_hours - total_completed, 0)
    completion = 0 if total_planned_hours <= 0 else min(total_completed / total_planned_hours * 100, 100)
    totals_by_field = {item["Field"]: item["Total"] for item in totals_from_daily_rows(rows)}
    bars = []
    for category, condition, percent in BURNDOWN_ROWS:
        field = PROGRESS_FIELD_BY_CONDITION.get(condition, condition)
        completed_text = totals_by_field.get(field, "00:00:00")
        completed_hours = parse_duration_to_hours(completed_text)
        planned_hours = total_planned_hours * percent / 100
        width = 0 if planned_hours <= 0 else min(completed_hours / planned_hours * 100, 100)
        bars.append(
            f"""
            <div class="chart-row">
              <div class="chart-label"><strong>{html.escape(condition)}</strong><span>{html.escape(category)}</span></div>
              <div class="bar-track"><div class="bar-fill" style="width: {width:.1f}%"></div></div>
              <div class="chart-number">{completed_hours:.1f} / {planned_hours:.1f} hrs</div>
            </div>
            """
        )
    return f"""
    <div class="chart-summary">
      <div class="donut" style="--done: {completion:.1f}%"><span>{completion:.1f}%</span></div>
      <div class="chart-stats">
        <div><strong>{total_planned_hours:.1f}</strong><span>Planned hours</span></div>
        <div><strong>{total_completed:.1f}</strong><span>Completed hours</span></div>
        <div><strong>{total_remaining:.1f}</strong><span>Remaining hours</span></div>
      </div>
    </div>
    <div class="chart-legend"><span class="done-dot"></span>Completed <span class="left-dot"></span>Remaining</div>
    <div class="bar-list">{''.join(bars)}</div>
    """


def csv_files_table_html(rows, vehicle=None, selected_source="", selected_action=""):
    if vehicle is not None:
        rows = [row for row in rows if str(row.get("Vehicle") or "Default").strip() == vehicle]
    if not rows:
        vehicle_text = f" for {html.escape(vehicle)}" if vehicle else ""
        return f"<p class='empty'>No CSV files uploaded{vehicle_text} yet.</p>"
    sources = []
    selected = ""
    rows_by_source = {}
    for row in rows:
        source = str(row.get("Source File") or "").strip()
        if not source:
            continue
        rows_by_source.setdefault(source, []).append(row)
        if source not in sources:
            sources.append(source)
    selected = selected_source if selected_source in rows_by_source else ""
    buttons = "".join(
        f"<a class='csv-select {'active' if source == selected else ''}' href='/?tab=csv-list&vehicle={quote_plus(vehicle or active_vehicle())}&source={quote_plus(source)}'>{index}. {html.escape(source)}</a>"
        for index, source in enumerate(sources, start=1)
    )
    detail_sections = []
    headers = [column[0] for column in DAILY_TRACKER_COLUMNS]
    for source in ([selected] if selected else []):
        source_rows = rows_by_source[source]
        body_rows = [
            "<tr>" + "".join(f"<td>{html.escape(str(row.get(header, '')))}</td>" for header in headers) + "</tr>"
            for row in source_rows
        ]
        action_bar = f"""
          <div class="csv-actions">
            <a class="button" href="/?tab=csv-list&vehicle={vehicle_param(vehicle)}&source={quote_plus(source)}&csv_action=edit">Edit</a>
            <a class="button" href="/download-uploaded-csv?vehicle={vehicle_param(vehicle)}&source={quote_plus(source)}">Download</a>
            <form action="/delete-csv?vehicle={vehicle_param(vehicle)}" method="post">
              <input type="hidden" name="source_file" value="{html.escape(source)}">
              <button class="danger" type="submit">Delete</button>
            </form>
            <a class="button secondary" href="/?tab=csv-list&vehicle={vehicle_param(vehicle)}">Close</a>
          </div>
        """
        detail_body = (
            tracker_table_html(source_rows, vehicle, return_tab="csv-list", source=source)
            if selected_action == "edit"
            else f"<table><thead><tr>{''.join(f'<th>{html.escape(header)}</th>' for header in headers)}</tr></thead><tbody>{''.join(body_rows)}</tbody></table>"
        )
        detail_sections.append(
            f"""
            <div class="csv-detail active" data-source="{html.escape(source)}">
              <h3>{html.escape(source)}</h3>
              {action_bar}
              <div class="muted">Download returns the original uploaded CSV when available. Older files uploaded before permanent CSV storage are downloaded as reconstructed tracker-row CSVs.</div>
              <div class="table-wrap">{detail_body}</div>
            </div>
            """
        )
    empty_detail = "" if selected else "<p class='empty'>Select a CSV file to view actions and data.</p>"
    return f"<div class='csv-file-layout'><div class='csv-file-list'>{buttons}</div><div class='csv-file-detail'>{empty_detail}{''.join(detail_sections)}</div></div>"


def seconds_to_hms(seconds):
    try:
        seconds = int(round(float(seconds or 0)))
    except ValueError:
        seconds = 0
    hours, remainder = divmod(seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"


def row_from_form(form):
    row = {}
    for header, _, _ in DAILY_TRACKER_COLUMNS:
        row[header] = form.get(header, [""])[0].strip()
    return row


def source_delete_form_html(vehicle=None):
    vehicle = normalize_vehicle(vehicle or active_vehicle())
    sources = source_files_from_database(OUTPUT_DIR, vehicle=vehicle)
    if not sources:
        return """
    <form class="delete-csv-form">
      <div class="source-list"><div class="muted">No uploaded CSV files found yet.</div></div>
      <button class="danger" type="button" disabled>Delete CSV Files</button>
    </form>"""
    options = "".join(
        "<label class='check-row'>"
        f"<input type='checkbox' name='source_file' value='{html.escape(source)}'>"
        f"<span>{html.escape(source)}</span>"
        "</label>"
        for source in sources
    )
    return f"""
    <form action="/delete-csv?vehicle={vehicle_param(vehicle)}" method="post" class="delete-csv-form">
      <div class="source-list">{options}</div>
      <button class="danger" type="submit">Delete CSV Files</button>
    </form>"""


def vehicle_selector_html(vehicle=None):
    current = normalize_vehicle(vehicle or active_vehicle())
    vehicles = vehicles_from_database(OUTPUT_DIR)
    if current not in vehicles:
        vehicles.insert(0, current)
    options = "".join(
        f"<option value='{html.escape(vehicle)}' {'selected' if vehicle == current else ''}>{html.escape(vehicle)}</option>"
        for vehicle in vehicles
    )
    active_sources = source_files_from_database(OUTPUT_DIR, vehicle=current)
    active_row_count = count_rows_in_database(OUTPUT_DIR, vehicle=current)
    return f"""
    <div class="vehicle-bar">
      <form action="/vehicle/select" method="post" id="vehicle-select-form">
        <label>Vehicle
          <select name="vehicle" onchange="this.form.submit()">{options}</select>
        </label>
        <button type="submit">Switch Vehicle</button>
      </form>
      <form action="/vehicle/add" method="post">
        <label>Add vehicle
          <input type="text" name="vehicle" placeholder="Vehicle name">
        </label>
        <button type="submit">Add Vehicle</button>
      </form>
      <form action="/vehicle/remove" method="post">
        <input type="hidden" name="vehicle" value="{html.escape(current)}">
        <button class="danger" type="submit" {'disabled' if current == 'Default' else ''}>Remove Vehicle</button>
      </form>
      <div class="vehicle-status">
        <strong>Active vehicle: {html.escape(current)}</strong>
        <span>{len(active_sources)} CSV file(s)</span>
        <span>{active_row_count} tracker row(s)</span>
        <span>{html.escape(tracker_path_for_vehicle(current).name)}</span>
      </div>
    </div>"""


def page(message="", processed=None, skipped=None, pending_folder="", active_tab="home", selected_source="", selected_csv_action="", vehicle=None):
    vehicle = normalize_vehicle(vehicle or active_vehicle())
    processed = processed or []
    skipped = skipped or []
    tracker_rows = current_rows(vehicle)
    planned_hours = current_planned_hours(vehicle)
    active_tab = active_tab if active_tab in {"home", "excel-editor", "csv-list", "progress", "charts"} else "home"
    skipped_cards = "".join(
        f"<li><strong>{html.escape(item.get('file', ''))}</strong> {html.escape(item.get('message', ''))}</li>"
        for item in skipped
    )
    process_anyway = ""
    if skipped_cards:
        if pending_folder:
            process_anyway = f"""
      <form action="/folder" method="post" class="inline-action">
        <input type="hidden" name="folder" value="{html.escape(pending_folder)}">
        <input type="hidden" name="allow_partial" value="1">
        <button type="submit">Process Anyway</button>
      </form>"""
        else:
            process_anyway = """
      <form action="/upload-anyway" method="post" class="inline-action">
        <button type="submit">Process Anyway</button>
      </form>"""
    notice = f"<div class='notice'>{html.escape(message)}</div>" if message else ""
    db_warning = ""
    fallback_reason = database_fallback_reason()
    if fallback_reason:
        db_warning = (
            "<div class='notice'><strong>Temporary local storage is active.</strong> "
            "The cloud database connection failed, so the app is open using SQLite fallback. "
            "On Render, redeploy using the Blueprint from GitHub or remove the manually entered DATABASE_URL so "
            "Render can inject DATABASE_URL from ride-report-db. "
            f"<span class='muted'>{html.escape(fallback_reason)}</span></div>"
        )
    skipped_notice = (
        f"<div class='notice'><strong>Skipped files</strong><ul class='skipped'>{skipped_cards}</ul>"
        f"<div class='muted'>Some expected fields are missing. You can still process the file and missing fields will stay blank.</div>{process_anyway}</div>"
        if skipped_cards
        else ""
    )
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>RideReport Daily Tracker</title>
  <style>
    :root {{
      color-scheme: light;
      --ink: #17212b;
      --muted: #5b6570;
      --line: #d8dee6;
      --panel: #ffffff;
      --accent: #166d5c;
      --accent-dark: #0f5044;
      --wash: #f4f7f9;
      --warn: #f7e3a1;
      --danger: #b42318;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "Segoe UI", Arial, sans-serif;
      background: var(--wash);
      color: var(--ink);
    }}
    h1 {{ margin: 0; font-size: 24px; letter-spacing: 0; }}
    main {{ padding: 22px 28px 36px; max-width: 1500px; margin: 0 auto; }}
    .tabs {{
      display: flex;
      gap: 8px;
      margin-bottom: 14px;
      border-bottom: 1px solid var(--line);
    }}
    .tab-button {{
      border: 1px solid var(--line);
      border-bottom: 0;
      border-radius: 8px 8px 0 0;
      background: #fff;
      color: var(--ink);
      padding: 11px 16px;
    }}
    .tab-button.active {{
      background: var(--accent);
      color: #fff;
      border-color: var(--accent);
    }}
    .tab-panel {{ display: none; }}
    .tab-panel.active {{ display: block; }}
    .panel {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 16px;
    }}
    h2 {{ margin: 0 0 12px; font-size: 16px; }}
    form {{ display: flex; gap: 10px; align-items: center; flex-wrap: wrap; }}
    .inline-action {{ margin-top: 10px; }}
    .editor-actions {{ display: flex; gap: 10px; align-items: center; flex-wrap: wrap; margin-bottom: 10px; }}
    .editor-actions .inline-action {{ margin-top: 0; }}
    .editor-table {{ max-height: 580px; }}
    .add-row-title {{ margin-top: 18px; }}
    input[type=file], input[type=text], td input, .add-grid input {{
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 9px;
      background: #fff;
      min-width: 280px;
    }}
    td input.field-drive-id {{ min-width: 310px; font-family: Consolas, monospace; }}
    select {{
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 9px;
      background: #fff;
      min-width: 180px;
    }}
    .vehicle-bar {{
      display: flex;
      gap: 12px;
      align-items: center;
      flex-wrap: wrap;
      margin-bottom: 14px;
    }}
    .vehicle-bar form {{ background: #fff; border: 1px solid var(--line); border-radius: 8px; padding: 10px; }}
    .vehicle-bar label {{ display: flex; gap: 8px; align-items: center; color: var(--muted); font-size: 13px; }}
    .vehicle-status {{
      display: flex;
      gap: 10px;
      align-items: center;
      flex-wrap: wrap;
      padding: 10px 12px;
      border-radius: 8px;
      background: #eaf6f2;
      border: 1px solid #9fd0c1;
    }}
    .vehicle-status strong {{ color: var(--accent-dark); }}
    .vehicle-status span {{
      background: #fff;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 4px 8px;
      font-size: 12px;
    }}
    .active-vehicle-banner {{
      margin-bottom: 14px;
      padding: 12px 14px;
      background: #102a43;
      color: #fff;
      border-radius: 8px;
      font-weight: 700;
    }}
    td input {{ min-width: 110px; width: 100%; padding: 6px; }}
    .actions {{ min-width: 150px; }}
    .actions form {{ display: inline-flex; margin: 0 4px 4px 0; }}
    .danger {{ background: var(--danger); }}
    .danger:hover {{ background: #7a1a12; }}
    .secondary {{ background: #5b6570; }}
    .secondary:hover {{ background: #424a53; }}
    .add-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
      gap: 10px;
      align-items: end;
    }}
    .add-grid label {{ display: grid; gap: 4px; color: var(--muted); font-size: 12px; }}
    .add-grid input {{ min-width: 0; width: 100%; }}
    button, .button {{
      appearance: none;
      border: 0;
      border-radius: 6px;
      background: var(--accent);
      color: #fff;
      padding: 10px 14px;
      font-weight: 700;
      cursor: pointer;
      text-decoration: none;
      display: inline-block;
    }}
    button:hover, .button:hover {{ background: var(--accent-dark); }}
    .downloads {{ display: flex; gap: 10px; flex-wrap: wrap; margin-top: 10px; }}
    .top-actions {{
      display: grid;
      grid-template-columns: minmax(280px, 1fr) minmax(320px, 1fr) auto;
      gap: 12px;
      align-items: start;
    }}
    .source-list {{
      max-height: 120px;
      overflow: auto;
      border: 1px solid var(--line);
      border-radius: 6px;
      background: #fff;
      padding: 8px;
      margin-bottom: 10px;
    }}
    .check-row {{
      display: flex;
      align-items: center;
      gap: 8px;
      padding: 4px 2px;
      font-size: 13px;
    }}
    .check-row input {{ min-width: auto; }}
    .delete-csv-form {{ display: block; }}
    .day-group {{
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 12px;
      margin-bottom: 14px;
      background: #fff;
    }}
    .day-group h3 {{ margin: 0 0 10px; font-size: 16px; }}
    .day-group h4 {{ margin: 0 0 8px; font-size: 13px; color: var(--muted); }}
    .day-columns {{
      display: grid;
      grid-template-columns: minmax(0, 1fr) minmax(0, 1fr);
      gap: 14px;
    }}
    .csv-file-layout {{
      display: grid;
      grid-template-columns: minmax(220px, 320px) minmax(0, 1fr);
      gap: 14px;
    }}
    .csv-file-list {{
      display: grid;
      gap: 8px;
      align-content: start;
    }}
    .csv-select {{
      background: #fff;
      color: var(--ink);
      border: 1px solid var(--line);
      text-align: left;
      font-weight: 600;
    }}
    .csv-select.active {{
      background: var(--accent);
      color: #fff;
      border-color: var(--accent);
    }}
    .csv-detail {{ display: none; }}
    .csv-detail.active {{ display: block; }}
    .csv-actions {{ display: flex; gap: 10px; flex-wrap: wrap; align-items: center; margin: 8px 0 12px; }}
    .csv-actions form {{ margin: 0; }}
    .chart-summary {{ display: grid; grid-template-columns: 220px minmax(0, 1fr); gap: 20px; align-items: center; }}
    .donut {{
      width: 190px;
      aspect-ratio: 1;
      border-radius: 50%;
      background: conic-gradient(#00a676 var(--done), #f0b429 0);
      display: grid;
      place-items: center;
      color: var(--ink);
      font-size: 26px;
      font-weight: 800;
      position: relative;
    }}
    .donut::after {{ content: ""; position: absolute; width: 118px; aspect-ratio: 1; border-radius: 50%; background: #fff; }}
    .donut span {{ position: relative; z-index: 1; }}
    .chart-stats {{ display: grid; grid-template-columns: repeat(3, minmax(120px, 1fr)); gap: 12px; }}
    .chart-stats div {{ border: 1px solid var(--line); border-radius: 8px; padding: 14px; background: #fff; }}
    .chart-stats strong {{ display: block; font-size: 24px; color: var(--accent-dark); }}
    .chart-stats span {{ color: var(--muted); font-size: 13px; }}
    .chart-legend {{ display: flex; gap: 14px; align-items: center; margin: 16px 0; color: var(--muted); }}
    .done-dot, .left-dot {{ width: 12px; height: 12px; display: inline-block; border-radius: 50%; margin-right: -8px; }}
    .done-dot {{ background: #00a676; }}
    .left-dot {{ background: #f0b429; }}
    .bar-list {{ display: grid; gap: 10px; }}
    .chart-row {{ display: grid; grid-template-columns: 220px minmax(180px, 1fr) 150px; gap: 12px; align-items: center; }}
    .chart-label span {{ display: block; color: var(--muted); font-size: 12px; }}
    .bar-track {{ height: 16px; border-radius: 999px; background: #f1e0a8; overflow: hidden; }}
    .bar-fill {{ height: 100%; border-radius: 999px; background: linear-gradient(90deg, #00a676, #3bb6d0); }}
    .chart-number {{ font-size: 12px; color: var(--muted); text-align: right; }}
    .notice {{
      margin: 18px 0;
      padding: 12px 14px;
      background: var(--warn);
      border-radius: 6px;
      border: 1px solid #e0c85f;
    }}
    .grid {{
      display: grid;
      grid-template-columns: minmax(0, 1fr) minmax(0, 1fr);
      gap: 16px;
      margin-top: 16px;
    }}
    .table-wrap {{
      overflow: auto;
      max-height: 520px;
      border: 1px solid var(--line);
      border-radius: 6px;
    }}
    table {{ width: 100%; border-collapse: collapse; font-size: 13px; background: #fff; }}
    th, td {{ padding: 8px 10px; border-bottom: 1px solid var(--line); text-align: left; white-space: nowrap; }}
    th {{ position: sticky; top: 0; background: #eef3f6; z-index: 1; }}
    .muted {{ color: var(--muted); font-size: 13px; margin-top: 8px; }}
    .empty {{ color: var(--muted); margin: 0; }}
    .skipped {{ color: var(--danger); }}
    ul {{ margin: 8px 0 0; padding-left: 18px; }}
    @media (max-width: 850px) {{
      main, header {{ padding-left: 16px; padding-right: 16px; }}
      input[type=file], input[type=text] {{ min-width: 100%; }}
      .top-actions {{ grid-template-columns: 1fr; }}
      .grid {{ grid-template-columns: 1fr; }}
      .day-columns {{ grid-template-columns: 1fr; }}
      .csv-file-layout {{ grid-template-columns: 1fr; }}
      .chart-summary {{ grid-template-columns: 1fr; }}
      .chart-stats {{ grid-template-columns: 1fr; }}
      .chart-row {{ grid-template-columns: 1fr; }}
      .chart-number {{ text-align: left; }}
    }}
  </style>
</head>
<body>
  <main>
    {vehicle_selector_html(vehicle)}
    <div class="active-vehicle-banner">Showing data only for: {html.escape(vehicle)}</div>
    <nav class="tabs">
      <a class="tab-button {'active' if active_tab == 'home' else ''}" href="/?tab=home&vehicle={vehicle_param(vehicle)}">Home</a>
      <a class="tab-button {'active' if active_tab == 'csv-list' else ''}" href="/?tab=csv-list&vehicle={vehicle_param(vehicle)}">CSV Files</a>
      <a class="tab-button {'active' if active_tab == 'progress' else ''}" href="/?tab=progress&vehicle={vehicle_param(vehicle)}">Progress</a>
      <a class="tab-button {'active' if active_tab == 'charts' else ''}" href="/?tab=charts&vehicle={vehicle_param(vehicle)}">Charts</a>
    </nav>
    {db_warning}
    {notice}
    {skipped_notice}
    <section id="home" class="tab-panel {'active' if active_tab == 'home' else ''}">
      <div class="panel top-actions">
        <div>
          <form action="/upload?vehicle={vehicle_param(vehicle)}" method="post" enctype="multipart/form-data">
            <input type="file" name="files" accept=".csv" multiple>
            <button type="submit">Upload CSV Files to {html.escape(vehicle)}</button>
          </form>
        </div>
        <div>
          {source_delete_form_html(vehicle)}
        </div>
        <div>
          <a class="button" href="/?tab=excel-editor&vehicle={vehicle_param(vehicle)}">Edit Excel File</a>
        </div>
      </div>
    </section>
    <section id="excel-editor" class="tab-panel {'active' if active_tab == 'excel-editor' else ''}">
      <div class="panel">
        <h2>Edit Excel File - {html.escape(vehicle)}</h2>
        {browser_editor_html(tracker_rows, vehicle)}
      </div>
    </section>
    <section id="csv-list" class="tab-panel {'active' if active_tab == 'csv-list' else ''}">
      <div class="panel">
        <h2>CSV files list - {html.escape(vehicle)}</h2>
        <div class="muted">Only CSV files uploaded for the selected vehicle are shown here.</div>
        <div class="downloads"><a class="button" href="/download-uploaded-csv-list?vehicle={vehicle_param(vehicle)}">Download Uploaded CSV List</a></div>
        <div class="table-wrap">{csv_files_table_html(tracker_rows, vehicle, selected_source, selected_csv_action)}</div>
      </div>
    </section>
    <section id="progress" class="tab-panel {'active' if active_tab == 'progress' else ''}">
      <div class="panel">
        <h2>Progress - {html.escape(vehicle)}</h2>
        <form action="/settings/planned-hours?vehicle={vehicle_param(vehicle)}" method="post" class="inline-action">
          <label>Total planned hours
            <input type="text" name="total_planned_hours" value="{planned_hours:g}">
          </label>
          <button type="submit">Update Planned Hours</button>
        </form>
        <div class="muted">Category percentages remain fixed. Planned and remaining hours recalculate from this value.</div>
        <div class="table-wrap">{progress_table_html(tracker_rows, planned_hours)}</div>
      </div>
    </section>
    <section id="charts" class="tab-panel {'active' if active_tab == 'charts' else ''}">
      <div class="panel">
        <h2>Graphical Progress - {html.escape(vehicle)}</h2>
        {charts_html(tracker_rows, planned_hours, vehicle)}
      </div>
    </section>
  </main>
  <script>
    const buttons = document.querySelectorAll('.tab-button');
    const panels = document.querySelectorAll('.tab-panel');
    buttons.forEach((button) => {{
      button.addEventListener('click', () => {{
        buttons.forEach((item) => item.classList.remove('active'));
        panels.forEach((item) => item.classList.remove('active'));
        button.classList.add('active');
        document.getElementById(button.dataset.tab).classList.add('active');
      }});
    }});
  </script>
</body>
</html>"""


class Handler(BaseHTTPRequestHandler):
    def respond_html(self, content, status=200):
        data = content.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        parsed = urlparse(self.path)
        query = parse_qs(parsed.query)
        vehicle = normalize_vehicle(query.get("vehicle", [active_vehicle()])[0])
        active_tab = query.get("tab", ["home"])[0]
        selected_source = query.get("source", [""])[0]
        selected_csv_action = query.get("csv_action", [""])[0]
        if parsed.path == "/download":
            target = tracker_path_for_vehicle(vehicle)
            rebuild_tracker_from_database(OUTPUT_DIR, tracker_name=target.name, vehicle=vehicle)
            if not target.exists():
                self.send_error(404, "File not found")
                return
            data = target.read_bytes()
            filename = dated_excel_filename(target)
            self.send_response(200)
            self.send_header("Content-Type", mimetypes.guess_type(target.name)[0] or "application/octet-stream")
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if parsed.path == "/download-uploaded-csv-list":
            data = uploaded_csv_list_bytes(vehicle)
            filename = f"uploaded_csv_files_{vehicle}.csv"
            self.send_response(200)
            self.send_header("Content-Type", "text/csv; charset=utf-8")
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if parsed.path == "/download-uploaded-csv":
            source = query.get("source", [""])[0]
            stored = load_uploaded_csv_file(OUTPUT_DIR, vehicle, source)
            if not stored:
                stored = reconstructed_csv_from_rows(OUTPUT_DIR, vehicle, source)
            if not stored:
                self.send_error(404, "Uploaded CSV not found")
                return
            filename, data = stored
            self.send_response(200)
            self.send_header("Content-Type", "text/csv; charset=utf-8")
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if parsed.path == "/open-excel":
            target = tracker_path_for_vehicle(vehicle)
            rebuild_tracker_from_database(OUTPUT_DIR, tracker_name=target.name, vehicle=vehicle)
            if not target.exists():
                self.send_error(404, "File not found")
                return
            if not can_open_excel_locally():
                data = target.read_bytes()
                filename = dated_excel_filename(target)
                self.send_response(200)
                self.send_header("Content-Type", mimetypes.guess_type(target.name)[0] or "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
                return
            try:
                os.startfile(target)
                self.respond_html(page("Opened Daily Tracker in Excel.", vehicle=vehicle))
            except Exception as exc:
                self.respond_html(page(f"Could not open Excel file: {exc}", vehicle=vehicle), status=500)
            return
        self.respond_html(page(active_tab=active_tab, selected_source=selected_source, selected_csv_action=selected_csv_action, vehicle=vehicle))

    def do_POST(self):
        parsed = urlparse(self.path)
        query = parse_qs(parsed.query)
        vehicle = normalize_vehicle(query.get("vehicle", [active_vehicle()])[0])
        request_path = parsed.path
        length = int(self.headers.get("Content-Length", "0"))
        body = self.rfile.read(length)
        selected_source = ""
        selected_csv_action = ""
        try:
            if request_path == "/upload":
                csv_paths = parse_multipart(body, self.headers.get("Content-Type", ""))
                result = process_csv_paths(csv_paths, vehicle=vehicle)
                processed = result["processed"]
                skipped = result["skipped"]
                message = result["message"]
                pending_folder = ""
                active_tab = "home"
            elif request_path == "/upload-anyway":
                staging = UPLOAD_DIR / "_last_batch"
                result = process_reports(staging, OUTPUT_DIR, tracker_name=tracker_path_for_vehicle(vehicle).name, allow_partial=True, vehicle=vehicle)
                processed = result["processed"]
                skipped = result["skipped"]
                message = "Processed with missing fields left blank. " + result["message"]
                pending_folder = ""
                active_tab = "home"
            elif request_path == "/folder":
                form = parse_qs(body.decode("utf-8", "ignore"))
                folder = Path(form.get("folder", [str(DEFAULT_DOWNLOADS)])[0])
                allow_partial = form.get("allow_partial", ["0"])[0] == "1"
                result = process_reports(folder, OUTPUT_DIR, tracker_name=tracker_path_for_vehicle(vehicle).name, allow_partial=allow_partial, vehicle=vehicle)
                processed = result["processed"]
                skipped = result["skipped"]
                prefix = "Processed with missing fields left blank. " if allow_partial else ""
                message = prefix + result["message"] + f" Source: {folder}."
                pending_folder = str(folder)
                active_tab = "home"
            elif request_path == "/row/update":
                form = parse_qs(body.decode("utf-8", "ignore"))
                row_id = form.get("row_id", [""])[0]
                row = row_from_form(form)
                row["Vehicle"] = vehicle
                update_database_row(OUTPUT_DIR, row_id, row)
                refresh_uploaded_csv_from_rows(OUTPUT_DIR, vehicle, row.get("Source File", ""))
                rebuild_tracker_from_database(OUTPUT_DIR, tracker_name=tracker_path_for_vehicle(vehicle).name, vehicle=vehicle)
                processed = []
                skipped = []
                message = "Row updated in the stored CSV data and the Excel tracker."
                pending_folder = ""
                active_tab = form.get("return_tab", ["excel-editor"])[0] or "excel-editor"
                selected_source = form.get("source", [""])[0]
                selected_csv_action = "edit" if active_tab == "csv-list" and selected_source else ""
            elif request_path == "/row/delete":
                form = parse_qs(body.decode("utf-8", "ignore"))
                row_id = form.get("row_id", [""])[0]
                delete_database_row(OUTPUT_DIR, row_id)
                rebuild_tracker_from_database(OUTPUT_DIR, tracker_name=tracker_path_for_vehicle(vehicle).name, vehicle=vehicle)
                processed = []
                skipped = []
                message = "Row deleted from the tracker database."
                pending_folder = ""
                active_tab = "excel-editor"
            elif request_path == "/row/add":
                form = parse_qs(body.decode("utf-8", "ignore"))
                row = row_from_form(form)
                row["Vehicle"] = vehicle
                add_database_row(OUTPUT_DIR, row)
                rebuild_tracker_from_database(OUTPUT_DIR, tracker_name=tracker_path_for_vehicle(vehicle).name, vehicle=vehicle)
                processed = []
                skipped = []
                message = "Row added to the tracker database."
                pending_folder = ""
                active_tab = "excel-editor"
            elif request_path == "/sync-excel":
                rows = import_tracker_workbook(OUTPUT_DIR, tracker_path_for_vehicle(vehicle), vehicle=vehicle)
                processed = []
                skipped = []
                message = f"Synced {len(rows)} Excel row(s) to the tracker database. If Excel is open locally, save and close it before syncing again to rewrite recalculated totals into the workbook."
                pending_folder = ""
                active_tab = "excel-editor"
            elif request_path == "/refresh-totals":
                processed = []
                skipped = []
                message = "Refresh Totals was removed because totals update automatically when you upload, edit, or delete CSV data."
                pending_folder = ""
                active_tab = "excel-editor"
            elif request_path == "/delete-csv":
                form = parse_qs(body.decode("utf-8", "ignore"))
                selected_sources = form.get("source_file", [])
                deleted = delete_database_rows_by_source_files(OUTPUT_DIR, selected_sources, vehicle=vehicle)
                rebuild_tracker_from_database(OUTPUT_DIR, tracker_name=tracker_path_for_vehicle(vehicle).name, vehicle=vehicle)
                processed = []
                skipped = []
                if selected_sources:
                    message = f"Deleted {deleted} Daily Tracker row(s) from selected CSV file(s). They will stay removed unless you upload those CSV files again."
                else:
                    message = "No CSV files selected for deletion."
                pending_folder = ""
                active_tab = "csv-list"
            elif request_path == "/settings/planned-hours":
                form = parse_qs(body.decode("utf-8", "ignore"))
                raw_value = form.get("total_planned_hours", [str(DEFAULT_TOTAL_PLANNED_HOURS)])[0].strip()
                try:
                    planned_hours = float(raw_value)
                    if planned_hours < 0:
                        raise ValueError
                    set_setting(OUTPUT_DIR, vehicle_setting_key(vehicle, "total_planned_hours"), planned_hours)
                    message = f"Total planned hours updated to {planned_hours:g}."
                except ValueError:
                    message = "Please enter a valid non-negative planned hours value."
                processed = []
                skipped = []
                pending_folder = ""
                active_tab = "progress"
            elif request_path == "/vehicle/select":
                form = parse_qs(body.decode("utf-8", "ignore"))
                vehicle = normalize_vehicle(form.get("vehicle", ["Default"])[0])
                rebuild_tracker_from_database(OUTPUT_DIR, tracker_name=tracker_path_for_vehicle(vehicle).name, vehicle=vehicle)
                processed = []
                skipped = []
                message = f"Switched to vehicle: {vehicle}."
                pending_folder = ""
                active_tab = "home"
            elif request_path == "/vehicle/add":
                form = parse_qs(body.decode("utf-8", "ignore"))
                vehicle = form.get("vehicle", [""])[0].strip()
                if vehicle:
                    vehicle = normalize_vehicle(vehicle)
                    add_vehicle(OUTPUT_DIR, vehicle)
                    set_setting(OUTPUT_DIR, vehicle_setting_key(vehicle, "total_planned_hours"), DEFAULT_TOTAL_PLANNED_HOURS)
                    rebuild_tracker_from_database(OUTPUT_DIR, tracker_name=tracker_path_for_vehicle(vehicle).name, vehicle=vehicle)
                    message = f"Added and switched to vehicle: {vehicle}."
                else:
                    message = "Please enter a vehicle name."
                processed = []
                skipped = []
                pending_folder = ""
                active_tab = "home"
            elif request_path == "/vehicle/remove":
                form = parse_qs(body.decode("utf-8", "ignore"))
                vehicle = normalize_vehicle(form.get("vehicle", [vehicle])[0])
                if vehicle and vehicle != "Default":
                    deleted = remove_vehicle(OUTPUT_DIR, vehicle)
                    vehicle = "Default"
                    rebuild_tracker_from_database(OUTPUT_DIR, tracker_name=tracker_path_for_vehicle(vehicle).name, vehicle=vehicle)
                    message = f"Removed vehicle '{vehicle}' and deleted {deleted} row(s). Switched back to Default."
                else:
                    message = "Default vehicle cannot be removed."
                processed = []
                skipped = []
                pending_folder = ""
                active_tab = "home"
            else:
                self.send_error(404, "Not found")
                return
            self.respond_html(
                page(
                    message,
                    processed,
                    skipped,
                    pending_folder=pending_folder,
                    active_tab=active_tab,
                    selected_source=selected_source,
                    selected_csv_action=selected_csv_action,
                    vehicle=vehicle,
                )
            )
        except Exception as exc:
            self.respond_html(page(f"Could not process reports: {exc}", vehicle=vehicle), status=500)

    def log_message(self, format, *args):
        print(format % args)


def main():
    parser = argparse.ArgumentParser(description="Run the RideReport local web app.")
    parser.add_argument("--open", action="store_true", help="Open the app in the default browser.")
    parser.add_argument("--open-excel", action="store_true", help="Open the Daily Tracker workbook in Excel.")
    parser.add_argument("--refresh", action="store_true", help="Refresh outputs from Downloads before starting.")
    args = parser.parse_args()

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    if args.refresh:
        result = process_reports(DEFAULT_DOWNLOADS, OUTPUT_DIR, tracker_name=active_tracker_path().name, vehicle=active_vehicle())
        print(result["message"])
    host = os.environ.get("HOST", "127.0.0.1")
    if os.environ.get("PORT"):
        host = os.environ.get("HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", "8765"))
    server = ThreadingHTTPServer((host, port), Handler)
    local_url = f"http://127.0.0.1:{port}"
    print(f"RideReport app running at http://{host}:{port}")
    if args.open:
        try:
            webbrowser.open(local_url)
        except Exception:
            pass
    if args.open_excel and active_tracker_path().exists():
        try:
            os.startfile(active_tracker_path())
        except Exception:
            pass
    server.serve_forever()


if __name__ == "__main__":
    main()
