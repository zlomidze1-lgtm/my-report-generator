from flask import Flask, render_template, request, jsonify, send_file
import json
import os
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from datetime import datetime
import io
import psycopg2
from psycopg2.extras import RealDictCursor

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")

# Vercel Postgres-თან კავშირის ფუნქცია
def get_db_connection():
    # Vercel ავტომატურად ამატებს POSTGRES_URL ცვლადს
    conn = psycopg2.connect(
        os.environ.get("POSTGRES_URL"),
        cursor_factory=RealDictCursor
    )
    return conn

# ცხრილის ავტომატური შექმნა (თუ ჯერ არ არსებობს)
def init_db():
    try:
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS records (
                id VARCHAR(100) PRIMARY KEY,
                data JSONB NOT NULL,
                created_at DATE NOT NULL
            );
        """)
        conn.commit()
        cur.close()
        conn.close()
    except Exception as e:
        print("DB Init Error:", e)

# აპლიკაციის ჩართვისას იქმნება ბაზა
init_db()

# ... (აქ რჩება თქვენი HEADERS, COL, LEADER_POSITIONS და generate_excel_from_records ფუნქცია) ...

@app.route("/api/records", methods=["GET"])
def api_get_records():
    try:
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("SELECT data, created_at FROM records ORDER BY created_at DESC;")
        rows = cur.fetchall()
        cur.close()
        conn.close()

        records = [r["data"] for r in rows]
        
        # 1 წელზე (365 დღეზე) ძველი ჩანაწერების შემოწმება
        has_old = False
        now = datetime.now().date()
        for r in rows:
            if (now - r["created_at"]).days >= 365:
                has_old = True
                break

        return jsonify({"records": records, "has_old_records": has_old})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/records/save", methods=["POST"])
def api_save_record():
    data = request.get_json(silent=True)
    if not data or "id" not in data:
        return jsonify({"error": "არასწორი მონაცემები"}), 400

    rec_id = str(data["id"])
    rec_date = data.get("date", datetime.now().strftime("%Y-%m-%d"))

    try:
        conn = get_db_connection()
        cur = conn.cursor()
        # თუ ID არსებობს - ანახლებს, თუ არა - ამატებს (UPSERT)
        cur.execute("""
            INSERT INTO records (id, data, created_at)
            VALUES (%s, %s, %s)
            ON CONFLICT (id) DO UPDATE SET data = EXCLUDED.data, created_at = EXCLUDED.created_at;
        """, (rec_id, json.dumps(data, ensure_ascii=False), rec_date))
        conn.commit()
        cur.close()
        conn.close()
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/records/delete/<path:record_id>", methods=["DELETE"])
def api_delete_record(record_id):
    try:
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("DELETE FROM records WHERE id = %s;", (str(record_id),))
        conn.commit()
        cur.close()
        conn.close()
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500
