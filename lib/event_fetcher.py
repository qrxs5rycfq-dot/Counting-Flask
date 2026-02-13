import datetime
import asyncpg
import logging
from typing import List, Dict, Optional

log = logging.getLogger("db_event_fetcher")


class EventFetcher:
    def __init__(self, dsn: str):
        self.dsn = dsn
        self.api_offline = False
        self.conn = None

    async def connect(self):
        if self.conn is None:
            self.conn = await asyncpg.connect(dsn=self.dsn)

    async def close(self):
        if self.conn:
            await self.conn.close()
            self.conn = None

    async def fetch_range(
        self,
        start: datetime.datetime,
        end: datetime.datetime,
        page: int = 1,
        per_page: int = 800,
        order: str = 'desc',
    ) -> Optional[List[Dict]]:
        await self.connect()
        offset = (page - 1) * per_page
        order_clause = 'DESC' if order.lower() == 'desc' else 'ASC'

        query = f"""
            SELECT pin, name, dept_name, dev_alias, event_point_name, event_time, update_time
            FROM acc_transaction
            WHERE event_time BETWEEN $1 AND $2
            ORDER BY event_time {order_clause}
            OFFSET $3 LIMIT $4
        """

        try:
            rows = await self.conn.fetch(query, start, end, offset, per_page)
            return [dict(row) for row in rows]
        except Exception as e:
            self.api_offline = True
            log.error(f"[EventFetcher] Error fetching page {page}: {e}")
            return None

    async def _fetch_all(
        self,
        start: datetime.datetime,
        end: datetime.datetime,
        per_page: int = 800,
        order: str = 'desc'
    ) -> List[Dict]:
        page = 1
        result = []
        while True:
            page_data = await self.fetch_range(start, end, page, per_page, order)
            if not page_data:  # handle None and empty list
                break
            result.extend(page_data)
            page += 1
        return result

    async def fetch_combined_events(self, order: str = 'desc') -> List[Dict]:
        # Ambil waktu lokal dari sistem dengan timezone aware
        now = datetime.datetime.now().astimezone()
        today = now.replace(hour=0, minute=0, second=0, microsecond=0)
        yesterday = today - datetime.timedelta(days=1)
        tomorrow = today + datetime.timedelta(days=1)

        # Ubah ke offset-naive karena asyncpg dan DB tidak pakai tzinfo
        yesterday_naive = yesterday.replace(tzinfo=None)
        today_naive = today.replace(tzinfo=None)
        tomorrow_naive = tomorrow.replace(tzinfo=None)

        await self.connect()

        events_yesterday = await self._fetch_all(yesterday_naive, today_naive, order=order)
        events_today = await self._fetch_all(today_naive, tomorrow_naive, order=order)

        await self.close()

        combined = events_yesterday + events_today

        # Sortir ulang sesuai order agar konsisten
        reverse = order.lower() == 'desc'
        combined.sort(key=lambda x: x['event_time'], reverse=reverse)

        log.info(f"[EventFetcher] Total events fetched (2 hari): {len(combined)}")
        return combined
