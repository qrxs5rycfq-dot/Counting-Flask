import os
import psycopg2
from psycopg2.extras import RealDictCursor

def get_conn():
    return psycopg2.connect(dsn=os.getenv("DATABASE_URL"))

def get_transaksi_filtered(pin, nama, dept, dari, ke, page=1, per_page=50):
    conn = get_conn()
    cur = conn.cursor(cursor_factory=RealDictCursor)

    query = "SELECT * FROM acc_firstin_lastout WHERE update_time BETWEEN %s AND %s"
    params = [dari, ke]

    if pin:
        query += " AND pin ILIKE %s"
        params.append(f"%{pin}%")
    if nama:
        query += " AND name ILIKE %s"
        params.append(f"%{nama}%")
    if dept:
        query += " AND dept_name ILIKE %s"
        params.append(f"%{dept}%")

    # Total count
    count_query = f"SELECT COUNT(*) FROM ({query}) AS sub"
    cur.execute(count_query, params)
    total = cur.fetchone()["count"]

    # Pagination
    offset = (page - 1) * per_page
    query += " ORDER BY first_in_time NULLS LAST LIMIT %s OFFSET %s"
    params += [per_page, offset]

    cur.execute(query, params)
    result = cur.fetchall()

    cur.close()
    conn.close()
    return result, total