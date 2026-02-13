import logging
import asyncpg
from lib.person_detail import PersonDetailFetcher

log = logging.getLogger("api_tracker")

EMPTY_SUMMARY = {
    "offline": True,
    "totalin": 0,
    "totalout": 0,
    "totalcur": 0,
    "data": []
}


class SummaryBuilder:
    def __init__(self, db_dsn: str):
        self.db_dsn = db_dsn
        self.person_fetcher = PersonDetailFetcher()

    async def _build_person_detail(self, conn, pin: str, data: dict) -> dict:
        """
        Ambil detail person dari DB jika bukan visitor,
        langsung gunakan data visitor jika label visitor.
        """
        if data.get("label") == "visitor":
            detail = {
                "pin": pin,
                "name": data["name"],
                "time": data["last_time"],
                "label": "visitor",
                "company": data.get("company"),
                "visit_reason": data.get("visit_reason"),
                "host": data.get("host"),
            }
        else:
            detail = await self.person_fetcher.get(conn, pin, data["last_time"], data["name"]) or {}

        # Tandai possibly stuck dan kirim ke warning jika ada
        if data.get("possibly_stuck"):
            detail["possibly_stuck"] = True

        return detail

    async def build(self, per_person: dict) -> dict:
        summary = {
            "offline": False,
            "totalin": 0,
            "totalout": 0,
            "totalcur": 0,
            "data": [],
            "warning": []
        }

        departments = {}

        try:
            async with asyncpg.create_pool(dsn=self.db_dsn) as pool:
                async with pool.acquire() as conn:
                    for pin, data in per_person.items():
                        dept = data.get("dept") or "UNKNOWN"
                        dept_data = departments.setdefault(dept, {
                            "dept": dept,
                            "in": 0,
                            "out": 0,
                            "cur": 0,
                            "person": {"data": []}
                        })

                        # Pakai logical_in, logical_out, current hasil dari process_events
                        in_count = data.get("logical_in", 0)
                        out_count = data.get("logical_out", 0)
                        current_count = data.get("current", 0)

                        summary["totalin"] += in_count
                        summary["totalout"] += out_count
                        dept_data["in"] += in_count
                        dept_data["out"] += out_count
                        dept_data["cur"] += current_count
                        summary["totalcur"] += current_count

                        if current_count > 0:  # status inside setara current > 0
                            detail = await self._build_person_detail(conn, pin, data)

                            if data.get("possibly_stuck"):
                                summary["warning"].append(detail)

                            dept_data["person"]["data"].append(detail)

            summary["data"] = list(departments.values())
            return summary

        except Exception as e:
            log.critical(f"[DB] Failed to build summary: {e}")
            return EMPTY_SUMMARY.copy()
