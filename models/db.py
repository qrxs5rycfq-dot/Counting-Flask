import os
import psycopg2
from psycopg2.extras import RealDictCursor

def get_conn():
    return psycopg2.connect(dsn=os.getenv("DATABASE_URL"))

def get_grafik_data(dari, ke):
    """
    Fetch events from previous-day-midnight through ke in ONE combined list.
    This matches EventFetcher.fetch_combined_events() which combines
    yesterday + today into a single list for EventProcessor.
    """
    from datetime import datetime as _dt, timedelta

    conn = get_conn()
    cur = conn.cursor(cursor_factory=RealDictCursor)

    # Parse 'dari' to compute the previous day start (same as EventFetcher)
    if isinstance(dari, str):
        try:
            dari_dt = _dt.strptime(dari.strip(), "%Y-%m-%d %H:%M:%S")
        except ValueError:
            dari_dt = _dt.strptime(dari.strip()[:10], "%Y-%m-%d")
    else:
        dari_dt = dari

    # Start from previous day midnight (matching EventFetcher.fetch_combined_events)
    prev_start = dari_dt.replace(hour=0, minute=0, second=0) - timedelta(days=1)

    # Single query: previous-day-midnight → ke (combined, like EventFetcher)
    cur.execute("""
        SELECT pin, name, dept_name, dev_alias, event_point_name, event_time
        FROM acc_transaction
        WHERE event_time BETWEEN %s AND %s
        ORDER BY event_time ASC
    """, (prev_start, ke))
    events = cur.fetchall()

    cur.close()
    conn.close()
    return events


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