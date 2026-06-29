import io
import os
import urllib.parse
from pathlib import Path
from flask import Flask, jsonify, send_file, render_template, abort, request
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

BASE_PATH = Path(__file__).parent / "SoFlo1543"

app = Flask(__name__)
app.config['TEMPLATES_AUTO_RELOAD'] = True


@app.after_request
def add_cors_headers(response):
    """Allow the Streamlit HTML iframe to call local API routes (images, records, export)."""
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Methods"] = "GET, OPTIONS"
    return response


def hyperlink_to_filename(target):
    """Extract and URL-decode a filename from a hyperlink target like ./image/foo%20bar.jpg"""
    if not target or "image/" not in target:
        return None
    return urllib.parse.unquote(target.split("image/")[-1])


def get_retrieval_folders():
    folders = []
    if not BASE_PATH.exists():
        return folders
    for item in sorted(BASE_PATH.iterdir()):
        if item.is_dir() and item.name.startswith("Entrance and Exit Retrieval"):
            folders.append(item.name)
    return folders


def get_excel_path(folder_name):
    folder = BASE_PATH / folder_name
    if not folder.exists():
        return None
    for f in folder.iterdir():
        if f.suffix.lower() == ".xlsx":
            return f
    return None


def get_image_folder(folder_name):
    return BASE_PATH / folder_name / "image"


def parse_excel(folder_name):
    excel_path = get_excel_path(folder_name)
    if not excel_path:
        return None, "Excel file not found"

    img_folder = get_image_folder(folder_name)

    # keep_links=True ensures hyperlinks on cells are loaded (default, but explicit)
    wb = openpyxl.load_workbook(str(excel_path), keep_links=True)
    ws = wb.active

    meta = {
        "title": ws.cell(1, 1).value or "",
        "sys": ws.cell(2, 1).value or "",
        "export_time": ws.cell(3, 1).value or "",
        "total": ws.cell(4, 1).value or "",
    }

    records = []
    # Data rows start at Excel row 7; header is row 6.
    # Thumbnail hyperlinks are in columns 19-22 (S-V).
    for excel_row in range(7, ws.max_row + 1):
        plate = ws.cell(excel_row, 1).value
        if not plate:
            continue

        # Read images from hyperlinks on the four Thumbnail columns (19-22).
        # This is definitive — set by the parking system for each specific record.
        images = []
        seen = set()
        for col in range(19, 23):
            cell = ws.cell(excel_row, col)
            if cell.hyperlink:
                fname = hyperlink_to_filename(cell.hyperlink.target)
                if fname and fname not in seen:
                    # Only include if the file actually exists on disk
                    if (img_folder / fname).exists():
                        seen.add(fname)
                        images.append(fname)

        def val(col):
            v = ws.cell(excel_row, col).value
            return str(v) if v is not None else ""

        enter_time  = val(4)
        depart_time = val(5)

        records.append({
            "plate":        val(1),
            "parking_lot":  val(2),
            "enter_exit":   val(3),
            "enter_time":   enter_time,
            "depart_time":  depart_time,
            "duration":     val(6),
            "gate":         val(7),
            "method":       val(8),
            "result":       val(9),
            "reason":       val(10),
            "vehicle_list": val(12),
            "owner":        val(13),
            "region":       val(14),
            "vehicle_type": val(15),
            "color":        val(16),
            "brand":        val(17),
            "images":       images,
        })

    return {"meta": meta, "records": records}, None


def build_journeys_py(records):
    """Python equivalent of the JS buildJourneys function.

    Groups records by plate, pairs each entry with its exit using the same
    matching logic as the frontend (exact enter_time match first, then
    nearest-subsequent fallback), and returns a list of journey dicts:
        {'plate': str, 'entry': rec|None, 'exit': rec|None}
    """
    by_plate = {}
    for r in records:
        p = r["plate"]
        if p not in by_plate:
            by_plate[p] = {"entries": [], "exits": []}
        if r.get("enter_exit") == "Entering":
            by_plate[p]["entries"].append(r)
        else:
            by_plate[p]["exits"].append(r)

    journeys = []
    for plate, grp in by_plate.items():
        entries = sorted(grp["entries"], key=lambda r: r.get("enter_time") or "")
        exits   = sorted(grp["exits"],   key=lambda r: r.get("enter_time") or "")

        used_exits = set()
        for entry in entries:
            # Definitive match: exit whose enter_time equals this entry's enter_time
            match_idx = next(
                (i for i, ex in enumerate(exits)
                 if i not in used_exits
                 and ex.get("enter_time") == entry.get("enter_time")),
                -1,
            )
            # Fallback: nearest exit whose enter_time is >= entry's enter_time
            if match_idx == -1:
                match_idx = next(
                    (i for i, ex in enumerate(exits)
                     if i not in used_exits
                     and (ex.get("enter_time") or "") >= (entry.get("enter_time") or "")),
                    -1,
                )
            if match_idx != -1:
                used_exits.add(match_idx)
                journeys.append({"plate": plate, "entry": entry, "exit": exits[match_idx]})
            else:
                journeys.append({"plate": plate, "entry": entry, "exit": None})

        for i, ex in enumerate(exits):
            if i not in used_exits:
                journeys.append({"plate": plate, "entry": None, "exit": ex})

    journeys.sort(
        key=lambda j: (j.get("entry") or j.get("exit") or {}).get("enter_time") or ""
    )
    return journeys


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/folders")
def api_folders():
    folders = get_retrieval_folders()
    result = []
    for name in folders:
        # Parse timestamp from folder name: Entrance and Exit Retrieval_YYYYMMDDHHMMSS
        ts = name.split("_")[-1] if "_" in name else ""
        label = name
        if len(ts) == 14:
            label = f"{ts[0:4]}-{ts[4:6]}-{ts[6:8]} {ts[8:10]}:{ts[10:12]}:{ts[12:14]}"
        result.append({"name": name, "label": label})
    return jsonify(result)


@app.route("/api/records/<path:folder_name>")
def api_records(folder_name):
    data, err = parse_excel(folder_name)
    if err:
        return jsonify({"error": err}), 404
    return jsonify(data)


@app.route("/api/all-records")
def api_all_records():
    """Merge records from every folder into one response."""
    all_records = []
    for folder_name in get_retrieval_folders():
        data, err = parse_excel(folder_name)
        if err or not data:
            continue
        # Tag each record with its source folder so image URLs work
        for r in data["records"]:
            r["folder"] = folder_name
        all_records.extend(data["records"])
    # Sort chronologically by entering time
    all_records.sort(key=lambda r: r.get("enter_time") or "")
    return jsonify({"records": all_records, "meta": {"title": "All Folders Combined"}})


@app.route("/images/<path:folder_name>/<path:filename>")
def serve_image(folder_name, filename):
    img_path = BASE_PATH / folder_name / "image" / filename
    if not img_path.exists():
        abort(404)
    return send_file(str(img_path))


@app.route("/api/export-unmatched")
def api_export_unmatched():
    """Generate an Excel workbook containing unmatched records with embedded images.

    Query parameters:
        folder            – folder name or '__ALL__'
        type              – 'entry-only' | 'exit-only' | 'both'  (default: 'both')
        include_no_images – '1' to include records with no images, '0' to skip them
    """
    folder = request.args.get("folder", "")
    export_type = request.args.get("type", "both")
    include_no_images = request.args.get("include_no_images", "1") == "1"

    if not folder:
        return jsonify({"error": "folder parameter is required"}), 400

    # ── Gather records ──────────────────────────────────────────────────────
    if folder == "__ALL__":
        records = []
        for fn in get_retrieval_folders():
            data, err = parse_excel(fn)
            if err or not data:
                continue
            for r in data["records"]:
                r["folder"] = fn
            records.extend(data["records"])
    else:
        data, err = parse_excel(folder)
        if err:
            return jsonify({"error": err}), 404
        records = data["records"]
        for r in records:
            r.setdefault("folder", folder)

    # ── Build journeys and select unmatched rows ────────────────────────────
    journeys = build_journeys_py(records)

    if export_type == "entry-only":
        rows_to_export = [
            {"rec": j["entry"], "unmatched_type": "Entry Without Exit"}
            for j in journeys if j["entry"] and not j["exit"]
        ]
    elif export_type == "exit-only":
        rows_to_export = [
            {"rec": j["exit"], "unmatched_type": "Exit Without Entry"}
            for j in journeys if not j["entry"] and j["exit"]
        ]
    else:  # both
        rows_to_export = [
            {"rec": j["entry"], "unmatched_type": "Entry Without Exit"}
            for j in journeys if j["entry"] and not j["exit"]
        ] + [
            {"rec": j["exit"], "unmatched_type": "Exit Without Entry"}
            for j in journeys if not j["entry"] and j["exit"]
        ]

    if not include_no_images:
        rows_to_export = [item for item in rows_to_export if item["rec"].get("images")]

    # ── Build Excel workbook ────────────────────────────────────────────────
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Unmatched Records"

    hdr_fill  = PatternFill("solid", fgColor="1A1D27")
    hdr_font  = Font(bold=True, color="E2E8F0")
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    headers = [
        "Plate", "Unmatched Type", "Enter/Exit",
        "Enter Time", "Depart Time", "Duration",
        "Gate", "Region", "Result", "Reason",
        "Image 1", "Image 2", "Image 3", "Image 4",
    ]
    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(1, col_idx)
        cell.fill  = hdr_fill
        cell.font  = hdr_font
        cell.alignment = hdr_align
    ws.row_dimensions[1].height = 30

    col_widths = [15, 22, 12, 22, 22, 12, 20, 14, 10, 40, 16, 16, 16, 16]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    IMG_W, IMG_H = 108, 78   # pixels for embedded thumbnails
    ROW_H_PT    = 62          # ~83 px; tall enough to show the thumbnail

    data_align = Alignment(vertical="center", wrap_text=True)

    for row_idx, item in enumerate(rows_to_export, start=2):
        rec = item["rec"]
        ws.append([
            rec.get("plate", ""),
            item["unmatched_type"],
            rec.get("enter_exit", ""),
            rec.get("enter_time", ""),
            rec.get("depart_time", ""),
            rec.get("duration", ""),
            rec.get("gate", ""),
            rec.get("region", ""),
            rec.get("result", ""),
            rec.get("reason", ""),
        ])
        ws.row_dimensions[row_idx].height = ROW_H_PT

        for col_idx in range(1, 11):
            ws.cell(row_idx, col_idx).alignment = data_align

        # Colour the "Unmatched Type" cell for quick visual scanning
        type_cell = ws.cell(row_idx, 2)
        if item["unmatched_type"] == "Entry Without Exit":
            type_cell.font = Font(color="F59E0B", bold=True)
        else:
            type_cell.font = Font(color="EF4444", bold=True)

        # Embed images
        rec_folder = rec.get("folder", folder if folder != "__ALL__" else "")
        img_folder = get_image_folder(rec_folder)
        for img_offset, fname in enumerate(rec.get("images", [])[:4]):
            img_path = img_folder / fname
            if not img_path.exists():
                continue
            try:
                xl_img        = XLImage(str(img_path))
                xl_img.width  = IMG_W
                xl_img.height = IMG_H
                col_letter = get_column_letter(11 + img_offset)  # K, L, M, N
                ws.add_image(xl_img, f"{col_letter}{row_idx}")
            except Exception:
                pass

    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    type_label    = {"entry-only": "entry_without_exit",
                     "exit-only":  "exit_without_entry"}.get(export_type, "all_unmatched")
    folder_label  = (folder[:30].replace(" ", "_") if folder != "__ALL__" else "all_folders")
    download_name = f"unmatched_{type_label}_{folder_label}.xlsx"

    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=download_name,
    )


if __name__ == "__main__":
    print("Starting Car Plate Viewer...")
    print(f"Data folder: {BASE_PATH}")
    folders = get_retrieval_folders()
    print(f"Found {len(folders)} retrieval folders")
    app.run(debug=False, port=5000, host="127.0.0.1")
