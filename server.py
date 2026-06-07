import json
import os
import io
from fastapi import FastAPI, UploadFile, Form, HTTPException, Header
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import psycopg2
import psycopg2.extras
import xlrd
import openpyxl

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

DATABASE_URL = os.environ.get("DATABASE_URL")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "vm2026")

JESPER_TIPS = {1:{"home":2,"away":0},2:{"home":1,"away":0},3:{"home":2,"away":0},4:{"home":2,"away":1},5:{"home":0,"away":2},6:{"home":2,"away":0},7:{"home":0,"away":2},8:{"home":0,"away":1},9:{"home":4,"away":0},10:{"home":2,"away":0},11:{"home":0,"away":1},12:{"home":2,"away":1},13:{"home":4,"away":0},14:{"home":2,"away":1},15:{"home":0,"away":2},16:{"home":2,"away":1},17:{"home":2,"away":1},18:{"home":0,"away":2},19:{"home":2,"away":0},20:{"home":2,"away":0},21:{"home":3,"away":0},22:{"home":2,"away":0},23:{"home":1,"away":0},24:{"home":0,"away":2},25:{"home":1,"away":0},26:{"home":2,"away":0},27:{"home":3,"away":0},28:{"home":2,"away":0},29:{"home":1,"away":0},30:{"home":0,"away":1},31:{"home":3,"away":0},32:{"home":1,"away":1},33:{"home":2,"away":0},34:{"home":2,"away":0},35:{"home":2,"away":0},36:{"home":0,"away":2},37:{"home":4,"away":0},38:{"home":2,"away":0},39:{"home":4,"away":0},40:{"home":0,"away":2},41:{"home":2,"away":0},42:{"home":3,"away":0},43:{"home":1,"away":1},44:{"home":1,"away":1},45:{"home":4,"away":0},46:{"home":3,"away":0},47:{"home":0,"away":3},48:{"home":2,"away":0},49:{"home":1,"away":1},50:{"home":2,"away":1},51:{"home":0,"away":3},52:{"home":3,"away":0},53:{"home":0,"away":2},54:{"home":1,"away":1},55:{"home":0,"away":4},56:{"home":0,"away":3},57:{"home":1,"away":1},58:{"home":0,"away":3},59:{"home":1,"away":2},60:{"home":1,"away":1},61:{"home":1,"away":2},62:{"home":3,"away":0},63:{"home":0,"away":1},64:{"home":0,"away":2},65:{"home":1,"away":1},66:{"home":0,"away":2},67:{"home":0,"away":3},68:{"home":2,"away":0},69:{"home":0,"away":2},70:{"home":1,"away":1},71:{"home":0,"away":1},72:{"home":0,"away":3}}


# ── Database ───────────────────────────────────────────────────────────────────

def get_db():
    return psycopg2.connect(DATABASE_URL)


def init_db():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS participants (
            id SERIAL PRIMARY KEY,
            name TEXT UNIQUE NOT NULL,
            tips JSONB NOT NULL,
            created_at TIMESTAMP DEFAULT NOW()
        )
    """)
    cur.execute("SELECT COUNT(*) FROM participants")
    if cur.fetchone()[0] == 0:
        tips = {str(k): v for k, v in JESPER_TIPS.items()}
        cur.execute(
            "INSERT INTO participants (name, tips) VALUES (%s, %s)",
            ("Jesper", json.dumps(tips))
        )
    conn.commit()
    cur.close()
    conn.close()


init_db()


def load_data():
    conn = get_db()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute("SELECT name, tips FROM participants ORDER BY id")
    rows = cur.fetchall()
    cur.close()
    conn.close()
    return [{"name": r["name"], "tips": r["tips"]} for r in rows]


# ── Excel parsing ──────────────────────────────────────────────────────────────

def to_num(v):
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def parse_tips_from_bytes(file_bytes: bytes, filename: str) -> dict:
    tips = {}

    if filename.lower().endswith(".xls"):
        rb = xlrd.open_workbook(file_contents=file_bytes)
        rs = rb.sheet_by_index(0)
        for r in range(6, 42):
            if r >= rs.nrows:
                break
            n1 = to_num(rs.cell_value(r, 0))
            if n1 is not None and 1 <= n1 <= 36:
                h = to_num(rs.cell_value(r, 6))
                a = to_num(rs.cell_value(r, 7))
                if h is not None and a is not None:
                    tips[str(int(n1))] = {"home": int(h), "away": int(a)}
            n2 = to_num(rs.cell_value(r, 13))
            if n2 is not None and 37 <= n2 <= 72:
                h = to_num(rs.cell_value(r, 19))
                a = to_num(rs.cell_value(r, 20))
                if h is not None and a is not None:
                    tips[str(int(n2))] = {"home": int(h), "away": int(a)}

    elif filename.lower().endswith((".xlsx", ".xlsm")):
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        ws = wb.active
        for r in range(7, 43):
            n1 = to_num(ws.cell(r, 1).value)
            if n1 is not None and 1 <= n1 <= 36:
                h = to_num(ws.cell(r, 7).value)
                a = to_num(ws.cell(r, 8).value)
                if h is not None and a is not None:
                    tips[str(int(n1))] = {"home": int(h), "away": int(a)}
            n2 = to_num(ws.cell(r, 14).value)
            if n2 is not None and 37 <= n2 <= 72:
                h = to_num(ws.cell(r, 20).value)
                a = to_num(ws.cell(r, 21).value)
                if h is not None and a is not None:
                    tips[str(int(n2))] = {"home": int(h), "away": int(a)}

    return tips


# ── Routes ────────────────────────────────────────────────────────────────────

@app.get("/")
def serve_html():
    return FileResponse("vm2026.html")


@app.get("/api/participants")
def get_participants():
    return load_data()


@app.post("/api/participants")
async def add_participant(name: str = Form(...), file: UploadFile = Form(...)):
    filename = file.filename or ""
    if not filename.lower().endswith((".xls", ".xlsx", ".xlsm")):
        raise HTTPException(400, "Endast .xls, .xlsx och .xlsm stöds")

    content = await file.read()
    tips = parse_tips_from_bytes(content, filename)

    if len(tips) == 0:
        raise HTTPException(400, "Hittade inga tips i filen. Kontrollera att mål är ifyllda i kolumn G/H (match 1-36) och T/U (match 37-72).")

    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT id FROM participants WHERE LOWER(name) = LOWER(%s)", (name,))
    existing = cur.fetchone()
    if existing:
        cur.execute(
            "UPDATE participants SET tips = %s WHERE LOWER(name) = LOWER(%s)",
            (json.dumps(tips), name)
        )
        conn.commit()
        cur.close()
        conn.close()
        return {"ok": True, "count": len(tips), "updated": True}

    cur.execute(
        "INSERT INTO participants (name, tips) VALUES (%s, %s)",
        (name, json.dumps(tips))
    )
    conn.commit()
    cur.close()
    conn.close()
    return {"ok": True, "count": len(tips), "updated": False}


@app.delete("/api/participants/{name}")
def delete_participant(name: str, x_admin_password: str = Header(default="")):
    if x_admin_password != ADMIN_PASSWORD:
        raise HTTPException(401, "Fel lösenord")
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT id FROM participants WHERE LOWER(name) = LOWER(%s)", (name,))
    row = cur.fetchone()
    if not row:
        cur.close()
        conn.close()
        raise HTTPException(404, "Deltagare hittades inte")
    if row[0] == 1:
        cur.close()
        conn.close()
        raise HTTPException(403, "Kan inte ta bort administratören")
    cur.execute("DELETE FROM participants WHERE id = %s", (row[0],))
    conn.commit()
    cur.close()
    conn.close()
    return {"ok": True}
