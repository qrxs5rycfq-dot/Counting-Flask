import os
import sys
import json
import requests
import psycopg2
from models.models import get_session, ZoneData

def get_conn():
    try:
        return psycopg2.connect(dsn=os.getenv("DATABASE_URL"))
    except Exception as e:
        print(f"[DB ERROR] Failed to connect: {e}")
        return None

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'jpg', 'jpeg', 'png'}

def get_departments():
    result = {}
    try:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("SELECT code, name FROM auth_department ORDER BY name")
        rows = cur.fetchall()
        for code, name in rows:
            result[code] = name
        cur.close()
        conn.close()
    except Exception as e:
        print(f"Error fetching departments: {e}")
    return result

def get_zone_data(zone):
    try:
        with get_session() as session:
            record = session.query(ZoneData).filter_by(zone=zone).first()
            return json.loads(record.data) if record else {"offline": True}
    except Exception as e:
        print(f"[ZoneData ERROR] {e}")
        return {"offline": True}
