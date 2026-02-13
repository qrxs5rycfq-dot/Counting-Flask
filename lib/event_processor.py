import datetime
import logging
from typing import Optional, List, Dict

log = logging.getLogger("api_tracker")


class EventProcessor:
    STUCK_TIMEOUT = 12 * 3600  # 12 jam
    DUPLICATE_IN_THRESHOLD = 2 * 3600  # 2 jam

    def __init__(self, in_devices: set[str], out_devices: set[str]):
        # Simpan semua device seperti di env
        self.in_devices = {d.strip().upper() for d in in_devices}
        self.out_devices = {d.strip().upper() for d in out_devices}

        # Buat dua kategori
        self.reader_in_devices = {d for d in self.in_devices if "-READER" in d}
        self.reader_out_devices = {d for d in self.out_devices if "-READER" in d}
        self.normal_in_devices = {d for d in self.in_devices if "-READER" not in d}
        self.normal_out_devices = {d for d in self.out_devices if "-READER" not in d}

        # 🧠 Jangan reset counter kalau sudah ada (repeat run)
        if not hasattr(self, "event_point_used_total"):
            self.event_point_used_total = 0

    @staticmethod
    def timestamp_from_str(time_str: str) -> Optional[int]:
        try:
            dt = datetime.datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S")
            return int(dt.timestamp())
        except Exception:
            return None

    def get_type_from_device(self, dev_name: str) -> Optional[str]:
        """Menentukan apakah device termasuk in atau out"""
        dev_upper = dev_name.strip().upper()
        if dev_upper in {d.replace("-READER", "") for d in self.in_devices}:
            return "in"
        if dev_upper in {d.replace("-READER", "") for d in self.out_devices}:
            return "out"
        return None

    def _prepare_prev_lookup(self, prev_events: List[dict]) -> Dict[str, List[dict]]:
        prev_lookup = {}
        for e in prev_events:
            pin = e.get("pin", "").strip()
            dev = str(e.get("dev_alias") or "").strip().upper()
            time_raw = e.get("event_time") or e.get("time")
            time_str = str(time_raw).strip() if time_raw else ""
            ts = self.timestamp_from_str(time_str)
            ev_type = self.get_type_from_device(dev)
            if not pin or not ts or not ev_type:
                continue

            dt = datetime.datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S")
            if dt.time() >= datetime.time(21, 0) or dt.time() <= datetime.time(12, 0):
                prev_lookup.setdefault(pin, []).append({
                    "type": ev_type,
                    "ts": ts,
                    "time": time_str
                })
        return prev_lookup

    async def process_events(
        self,
        events: List[dict],
        prev_events: Optional[List[dict]] = None
    ) -> Dict[str, dict]:

        per_person = {}
        unknown_devices = set()
        prev_lookup = self._prepare_prev_lookup(prev_events) if prev_events else {}

        for e in events:
            pin = e.get("pin", "").strip()
            name = e.get("name", "").strip()
            is_visitor = e.get("label") == "visitor"

            dev_alias = str(e.get("dev_alias") or e.get("device") or "").strip().upper()
            event_point_name = str(e.get("event_point_name") or "").strip().upper()
            dept = str(e.get("dept_name") or e.get("department") or ("TAMU" if is_visitor else "") or "").strip()

            # Default gunakan dev_alias
            dev = dev_alias
            used_from = "dev_alias"

            # 🔹 Jika di env ada -READER → cocokkan event_point_name tanpa -READER
            for reader_dev in (self.reader_in_devices | self.reader_out_devices):
                base_name = reader_dev.replace("-READER", "").strip().upper()
                if event_point_name == base_name:
                    dev = event_point_name
                    used_from = "event_point_name"
                    self.event_point_used_total += 1
                    break

            # timestamp
            time_raw = e.get("event_time") or e.get("time")
            time_str = str(time_raw).strip() if time_raw else ""
            ts = self.timestamp_from_str(time_str)
            if not all([dept, pin, dev, time_str]) or not ts:
                continue

            ev_type = self.get_type_from_device(dev)
            if not ev_type:
                unknown_devices.add(dev.lower())
                continue

            log.debug(f"[EventProcessor] {pin} → using {used_from}='{dev}' (type={ev_type})")

            person = per_person.setdefault(pin, {
                "dept": dept,
                "name": name,
                "events": [],
                "label": "visitor" if is_visitor else None,
                "company": e.get("company") if is_visitor else None,
                "visit_reason": e.get("visit_reason") if is_visitor else None,
                "host": e.get("host") if is_visitor else None,
            })
            person["events"].append({"type": ev_type, "ts": ts, "time": time_str})

        # 🔹 Gabungkan prev_lookup
        for pin, prev_evs in prev_lookup.items():
            if pin in per_person:
                events_list = per_person[pin]["events"]
                for prev_ev in prev_evs:
                    if prev_ev["type"] == "in" and not any(ev["ts"] == prev_ev["ts"] for ev in events_list):
                        events_list.append(prev_ev)
                per_person[pin]["events"] = sorted(events_list, key=lambda x: x["ts"])
            else:
                in_prev = [ev for ev in prev_evs if ev["type"] == "in"]
                if in_prev:
                    best_in = max(in_prev, key=lambda x: x["ts"])
                    per_person[pin] = {
                        "dept": "UNKNOWN",
                        "name": "",
                        "events": [best_in],
                        "label": None,
                        "company": None,
                        "visit_reason": None,
                        "host": None,
                    }

        # 🔹 Proses akhir
        result = {}
        for pin, person in per_person.items():
            events_sorted = sorted(person["events"], key=lambda x: x["ts"])
            status = "outside"
            last_time = ""
            filtered_events = []
            possibly_stuck = False
            last_ts = None
            logical_in = 0
            logical_out = 0
            current = 0

            for ev in events_sorted:
                ev_type = ev["type"]
                ts = ev["ts"]
                time_str = ev["time"]
                last_ts = ts

                if ev_type == "in" and status == "outside":
                    status = "inside"
                    logical_in += 1
                    current += 1
                    filtered_events.append(ev)
                    last_time = time_str
                elif ev_type == "out" and status == "inside":
                    status = "outside"
                    logical_out += 1
                    current -= 1
                    filtered_events.append(ev)
                    last_time = time_str

            if not filtered_events:
                continue

            if status == "inside" and last_ts is not None:
                last_in_event = next((e for e in reversed(filtered_events) if e["type"] == "in"), None)
                if last_in_event and (last_ts - last_in_event["ts"] > self.STUCK_TIMEOUT):
                    possibly_stuck = True

            person_result = {
                "dept": person["dept"],
                "name": person["name"],
                "status": status,
                "last_time": last_time,
                "events": filtered_events,
                "logical_in": logical_in,
                "logical_out": logical_out,
                "current": current,
            }

            if person.get("label") == "visitor":
                person_result["label"] = "visitor"
                if person.get("company"):
                    person_result["company"] = person["company"]
                if person.get("visit_reason"):
                    person_result["visit_reason"] = person["visit_reason"]
                if person.get("host"):
                    person_result["host"] = person["host"]

            if possibly_stuck:
                person_result["possibly_stuck"] = True

            result[pin] = person_result

        # 🔹 Logging
        log.info(f"[EventProcessor] Total processed people: {len(result)}")
        visitor_count = sum(1 for v in result.values() if v.get("label") == "visitor")
        log.info(f"[EventProcessor] Total visitors: {visitor_count}")
        log.info(f"[EventProcessor] Total event_point_name used: {self.event_point_used_total}")

        if unknown_devices:
            log.warning(f"[EventProcessor] Unknown devices: {', '.join(sorted(unknown_devices))}")

        sorted_result = dict(
            sorted(result.items(), key=lambda item: self.timestamp_from_str(item[1].get("last_time", "")) or 0, reverse=True)
        )
        return sorted_result
