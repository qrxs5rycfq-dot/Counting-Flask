import asyncpg
import logging
from datetime import datetime, time
from typing import List, Dict

log = logging.getLogger("api_tracker")


class VisitorFetcher:
    def __init__(self, db_dsn: str):
        self.db_dsn = db_dsn

    async def fetch_events(self) -> List[Dict]:
        today = datetime.now().date()
        start_time = datetime.combine(today, time.min)
        end_time = datetime.combine(today, time.max)

        try:
            async with asyncpg.create_pool(dsn=self.db_dsn) as pool:
                async with pool.acquire() as conn:
                    rows = await conn.fetch("""
                        SELECT
                            pin,
                            name,
                            dev_alias,
                            event_point_name,
                            event_time
                        FROM vis_visitor_lastaddr
                        WHERE event_time BETWEEN $1 AND $2
                    """, start_time, end_time)

            events = [
                {
                    "pin": str(row["pin"]),
                    "name": row["name"] or "TIDAK DIKETAHUI",
                    "dev_alias": row["dev_alias"] or "",
                    "event_point_name": row["event_point_name"] or "",
                    "time": row["event_time"],
                    "department": "VISITOR",
                    "label": "visitor"
                }
                for row in rows
            ]

            log.info(f"[VisitorFetcher] {len(events)} visitor events fetched.")
            return events

        except Exception as e:
            log.error(f"[VisitorFetcher] Error fetching visitor events: {e}")
            return []


async def enrich_visitor_details(db_dsn: str, events: List[Dict]) -> List[Dict]:
    if not events:
        return events

    pins = list({e["pin"] for e in events})
    if not pins:
        return events

    today = datetime.now().date()
    start_time = datetime.combine(today, time.min)
    end_time = datetime.combine(today, time.max)

    try:
        async with asyncpg.create_pool(dsn=db_dsn) as pool:
            async with pool.acquire() as conn:
                rows = await conn.fetch("""
                    SELECT
                        vis_emp_pin,
                        vis_company,
                        visit_reason,
                        visited_emp_dept,
                        visited_emp_name
                    FROM vis_transaction
                    WHERE vis_emp_pin = ANY($1)
                    AND update_time BETWEEN $2 AND $3
                """, pins, start_time, end_time)

        detail_map = {str(r["vis_emp_pin"]): r for r in rows}

        for event in events:
            detail = detail_map.get(event["pin"])
            if detail:
                event["company"] = detail["vis_company"]
                event["visit_reason"] = detail["visit_reason"]
                event["host"] = {
                    "name": detail["visited_emp_name"],
                    "department": detail["visited_emp_dept"]
                }

        log.info(f"[VisitorDetail] Enriched visitor details for {len(detail_map)} pins.")
        return events

    except Exception as e:
        log.error(f"[VisitorDetail] Error enriching visitor details: {e}")
        return events
