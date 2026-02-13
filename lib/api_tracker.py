import os
import sys
import logging
from dotenv import load_dotenv

from lib.event_fetcher import EventFetcher  
from lib.event_processor import EventProcessor
from lib.summary_builder import SummaryBuilder
from lib.visitor_fetcher import VisitorFetcher, enrich_visitor_details

# === Setup Environment ===
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
sys.path.append(BASE_DIR)
load_dotenv(os.path.join(BASE_DIR, ".env"))

# === Logging Setup ===
logging.basicConfig(level=logging.INFO, format="[%(asctime)s] [%(levelname)s] %(message)s")
log = logging.getLogger("api_tracker")

# === Default Summary ketika DB offline ===
EMPTY_SUMMARY = {
    "offline": True,
    "totalin": 0,
    "totalout": 0,
    "totalcur": 0,
    "data": []
}


class AsyncApiTracker:
    def __init__(self, in_devices=None, out_devices=None):
        self.in_devices = set(d.strip().lower() for d in in_devices or [])
        self.out_devices = set(d.strip().lower() for d in out_devices or [])
        self.db_dsn = os.getenv("DATABASE_URL")
        if not self.db_dsn:
            log.warning("[AsyncApiTracker] DATABASE_URL environment variable not set!")

        # === Inisialisasi semua komponen utama ===
        self.fetcher = EventFetcher(dsn=self.db_dsn)
        self.visitor_fetcher = VisitorFetcher(self.db_dsn)
        self.processor = EventProcessor(self.in_devices, self.out_devices)
        self.summary_builder = SummaryBuilder(self.db_dsn)

    async def run(self):
        try:
            # Ambil data kehadiran 2 hari terakhir, urut terbaru dulu
            events = await self.fetcher.fetch_combined_events(order='desc')
            if self.fetcher.api_offline:
                log.warning("[AsyncApiTracker] DB offline — returning EMPTY_SUMMARY")
                return EMPTY_SUMMARY.copy()

            # Ambil visitor events & enrich
            visitor_events = await self.visitor_fetcher.fetch_events()
            log.info(f"[AsyncApiTracker] Visitor events fetched: {len(visitor_events)}")

            visitor_events = await enrich_visitor_details(self.db_dsn, visitor_events)

            # Gabungkan semua events (karyawan + visitor)
            all_events = events + visitor_events
            log.info(f"[AsyncApiTracker] Total combined events (events + visitors): {len(all_events)}")

            # Proses semua events jadi per-person status
            per_person = await self.processor.process_events(all_events)

            # Bangun ringkasan akhir
            summary = await self.summary_builder.build(per_person)
            return summary

        except Exception as e:
            log.exception(f"[AsyncApiTracker] Unexpected error: {e}")
            return EMPTY_SUMMARY.copy()
