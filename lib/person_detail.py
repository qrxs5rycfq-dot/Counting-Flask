import os
import logging
from typing import Optional

log = logging.getLogger("api_tracker")


class PersonDetailFetcher:
    def __init__(self):
        self.cache = {}
        self.custom_keys = [
            k.strip().lower() for k in os.getenv("CUSTOM_ATTRIBUT", "").split(",") if k.strip()
        ]

    async def get(self, conn, pin: str, last_time: str, name: str) -> Optional[dict]:
        if pin in self.cache:
            cached = self.cache[pin].copy()
            cached["time"] = last_time
            cached["name"] = name or cached["name"]
            return cached

        try:
            # Ambil data dari pers_person
            person = await conn.fetchrow("""
                SELECT id, pin, name, gender
                FROM pers_person
                WHERE pin = $1
            """, pin)
            if not person:
                return {}

            person_id = person["id"]

            # Ambil plat dari park_car_number
            plat = ""
            car = await conn.fetchrow(
                "SELECT car_number FROM park_car_number WHERE person_id = $1", person_id
            )
            if car and car.get("car_number"):
                plat = car["car_number"]

            detail = {
                "name": name or person["name"],
                "id": person["pin"],
                "time": last_time,
                "gender": {"M": "Male", "F": "Female"}.get(person["gender"], ""),
                "plat": plat,
            }

            # Mapping NIPEG/JABATAN/KODE dari env
            attr_mapping = {}
            if self.custom_keys:
                attr_rows = await conn.fetch(
                    "SELECT attr_name, filed_index FROM pers_attribute WHERE LOWER(attr_name) = ANY($1)",
                    self.custom_keys,
                )
                attr_mapping = {r["attr_name"].lower(): r["filed_index"] for r in attr_rows}

            # Ambil dari pers_attribute_ext
            if attr_mapping:
                attr_ext = await conn.fetchrow(
                    "SELECT * FROM pers_attribute_ext WHERE person_id = $1", person_id
                )
                for key, idx in attr_mapping.items():
                    if attr_ext and idx < len(attr_ext):
                        detail[key] = attr_ext[idx]

            self.cache[pin] = detail.copy()
            return detail

        except Exception as e:
            log.error(f"[DB] Error getting detail for {pin}: {e}")
            return {}
