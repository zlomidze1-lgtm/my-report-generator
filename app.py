from flask import Flask, render_template, request, jsonify, send_file
import json
import os
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from datetime import datetime
import io

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")

HEADERS = [
    "ქალაქი (ლოკაცია)", "სახელი გვარი", "პირადი #", "ბრიგადის ნომერი", "თარიღი", "ID",
    "ობიექტის დასახელება", "მისამართი", "ქალაქი კოეფიციენტი (გადაადგილება)", "შესრულებული სამუშაო",
    "ბრიგადის წევრთა რაოდენობა", "დაწყება", "დასრულება", "რაოდენობა ჯამი",
    "რაოდენობა კაცზე", "კომენტარი", "შენიშვნა", "თანამდებობა"
]

COL = {name: i + 1 for i, name in enumerate([
    "city", "name", "position_id", "brigade", "date", "id",
    "object_name", "address", "coefficient", "work_type",
    "member_count", "start", "end", "qty_total",
    "qty_per_person", "comment", "note", "position"
])}

LEADER_POSITIONS = {"ბრიგადირი", "სარემონტო ბრიგადის უფროსი"}
NO_RECORD_TEXT = "არ არის ჩანაწერი"


def load_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def generate_excel_from_records(records):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "შესრულებული სამუშაოები"

    ws.append(HEADERS)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"

    row_num = 2

    for record in records:
        id_val = record.get("id", "")
        brigade = record.get("brigade", "")
        city = record.get("city", "")
        date = record.get("date", "")
        object_name = record.get("object_name", "")
        address = record.get("address", "")
        coefficient = record.get("coefficient", "")
        overall_comment = record.get("overall_comment", "")
        member_data = record.get("members", [])
        works = record.get("works", [])

        active_members, absent_members = [], []
        for m in member_data:
            note = (m.get("note") or "").strip()
            entry = {
                "name": m["name"],
                "position": m["position"],
                "personal_id": m.get("personal_id", ""),
                "extra": bool(m.get("extra")),
                "home_brigade": m.get("home_brigade"),
            }
            if m.get("absent", False) or note:
                entry["note"] = note or "არ გამოცხადება"
                absent_members.append(entry)
            else:
                active_members.append(entry)

        # ------------------ იერარქიული სორტირება ------------------
        # ბრიგადირი / სარემონტო ბრიგადის უფროსი ექცევა სიაში პირველ ადგილზე
        active_members.sort(key=lambda x: 0 if x.get("position") in LEADER_POSITIONS else 1)
        absent_members.sort(key=lambda x: 0 if x.get("position") in LEADER_POSITIONS else 1)
        # ------------------------------------------------------------

        worker_count = sum(1 for m in active_members if m["position"] not in LEADER_POSITIONS)

        def position_id_label(member):
            label = f"{member['position']} {member['personal_id']}".strip()
            if member.get("extra"):
                label += f" (ბრიგადა {member.get('home_brigade', '')}-დან)"
            return label

        def member_count_value(member):
            return 1 if member["position"] in LEADER_POSITIONS else worker_count

        for work in works:
            work_type = work.get("work_type", "")
            start_val = work.get("start", "")
            end_val = work.get("end", "")
            try:
                total_qty = float(work.get("quantity", 0) or 0)
            except (TypeError, ValueError):
                total_qty = 0
            work_comment = work.get("comment", "")

            for member in active_members:
                divisor = member_count_value(member)
                per_person = round(total_qty / divisor, 2) if divisor else 0
                row = [""] * len(HEADERS)
                row[COL["city"] - 1] = city
                row[COL["name"] - 1] = member["name"]
                row[COL["position_id"] - 1] = position_id_label(member)
                row[COL["brigade"] - 1] = brigade
                row[COL["date"] - 1] = date
                row[COL["id"] - 1] = id_val
                row[COL["object_name"] - 1] = object_name
                row[COL["address"] - 1] = address
                row[COL["coefficient"] - 1] = coefficient
                row[COL["work_type"] - 1] = work_type
                row[COL["member_count"] - 1] = member_count_value(member)
                row[COL["start"] - 1] = start_val
                row[COL["end"] - 1] = end_val
                row[COL["qty_total"] - 1] = total_qty
                row[COL["qty_per_person"] - 1] = per_person
                row[COL["comment"] - 1] = work_comment if work_comment else overall_comment
                row[COL["note"] - 1] = ""
                row[COL["position"] - 1] = member["position"]
                ws.append(row)
                row_num += 1

        for member in absent_members:
            row = [""] * len(HEADERS)
            row[COL["city"] - 1] = city
            row[COL["name"] - 1] = member["name"]
            row[COL["position_id"] - 1] = position_id_label(member)
            row[COL["brigade"] - 1] = brigade
            row[COL["date"] - 1] = date
            row[COL["id"] - 1] = id_val
            row[COL["object_name"] - 1] = object_name
            row[COL["address"] - 1] = address
            row[COL["coefficient"] - 1] = coefficient
            row[COL["work_type"] - 1] = "არ გამოცხადება"
            row[COL["member_count"] - 1] = 0
            row[COL["start"] - 1] = NO_RECORD_TEXT
            row[COL["end"] - 1] = NO_RECORD_TEXT
            row[COL["note"] - 1] = member["note"]
            row[COL["position"] - 1] = member["position"]
            ws.append(row)
            row_num += 1

    widths = [14, 22, 26, 12, 12, 10, 20, 20, 12, 30, 12, 10, 10, 12, 12, 20, 18, 22]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=len(HEADERS)):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}1"

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/get_brigades")
def get_brigades():
    config = load_config()
    brigades = sorted(config["brigades"].keys(), key=lambda x: int(x))
    return jsonify(brigades)


@app.route("/get_brigade_cities")
def get_brigade_cities():
    brigade = request.args.get("brigade")
    if not brigade:
        return jsonify([])
    config = load_config()
    members = config["brigades"].get(brigade, {}).get("members", [])
    cities = sorted({m["city"] for m in members if "city" in m})
    return jsonify(cities)


@app.route("/get_brigade_members")
def get_brigade_members():
    brigade = request.args.get("brigade")
    city = request.args.get("city", "")
    if not brigade:
        return jsonify([])
    config = load_config()
    members = config["brigades"].get(brigade, {}).get("members", [])
    if city:
        members = [m for m in members if m.get("city") == city]
    return jsonify(members)


@app.route("/get_all_members")
def get_all_members():
    config = load_config()
    all_members = []
    for brigade_num, data in config["brigades"].items():
        for m in data.get("members", []):
            all_members.append({
                "name": m["name"],
                "position": m["position"],
                "personal_id": m.get("personal_id", ""),
                "city": m.get("city", ""),
                "brigade": brigade_num,
            })
    all_members.sort(key=lambda m: (int(m["brigade"]), m["name"]))
    return jsonify(all_members)


@app.route("/get_work_types")
def get_work_types():
    config = load_config()
    return jsonify(config.get("work_types", []))


@app.route("/generate", methods=["POST"])
def generate():
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"error": "მონაცემები არასწორია"}), 400

    records = data.get("records", [])
    if not records:
        return jsonify({"error": "ჩანაწერები ვერ მოიძებნა"}), 400

    for rec in records:
        for field in ("id", "brigade", "city", "date", "address"):
            if not rec.get(field):
                return jsonify({"error": f"აკლია სავალდებულო ველი: {field}"}), 400
        if not rec.get("works"):
            return jsonify({"error": f"ჩანაწერს ID {rec.get('id')} არ აქვს სამუშაოები"}), 400

    try:
        output = generate_excel_from_records(records)
    except Exception as exc:
        return jsonify({"error": f"Excel-ის გენერირება ვერ მოხერხდა: {exc}"}), 500

    filename = f"ანგარიში_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        output,
        download_name=filename,
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


app = app

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
