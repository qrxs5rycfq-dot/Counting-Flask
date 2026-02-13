# tracker_worker.py

import asyncio
import json
import logging
import os
import platform
import signal
import sys
from typing import List

from dotenv import load_dotenv

# ─── Base Dir Detection ─────────────────────────────
def get_base_dir():
    if getattr(sys, 'frozen', False):  # PyInstaller
        return os.path.dirname(sys.executable)
    return os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

BASE_DIR = get_base_dir()
sys.path.append(BASE_DIR)
load_dotenv(os.path.join(BASE_DIR, ".env"))

# ─── Logging ────────────────────────────────────────
def setup_logging():
    log = logging.getLogger("tracker_worker")
    log.setLevel(logging.INFO)

    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S")
    
    if sys.stdout:
        handler = logging.StreamHandler(sys.stdout)
    else:
        os.makedirs("logs", exist_ok=True)
        handler = logging.FileHandler("logs/worker.log")

    handler.setFormatter(formatter)
    log.addHandler(handler)
    return log

log = setup_logging()

# ─── Import setelah path fix ────────────────────────
from models.models import ZoneData, get_session
from lib.api_tracker import AsyncApiTracker

# ─── Konfigurasi Zona ───────────────────────────────
ZONES: List[dict] = [
    {"name": "hijau", "in_env": "IN_DEVICES_HIJAU", "out_env": "OUT_DEVICES_HIJAU", "interval_env": "INTERVAL_HIJAU_SEC"},
    {"name": "merah", "in_env": "IN_DEVICES_MERAH", "out_env": "OUT_DEVICES_MERAH", "interval_env": "INTERVAL_MERAH_SEC"},
]

# ─── Worker Logic ───────────────────────────────────
async def fetch_and_store(zone: str, in_devices: list[str], out_devices: list[str]):
    try:
        log.info("[%s] Fetching...", zone.upper())
        tracker = AsyncApiTracker(in_devices, out_devices)
        data = await asyncio.wait_for(tracker.run(), timeout=120)

        if not isinstance(data, dict) or data.get("offline"):
            log.warning("[%s] Data invalid / offline", zone.upper())
            return

        with get_session() as session:
            session.query(ZoneData).filter_by(zone=zone).delete()
            session.add(ZoneData(zone=zone, data=json.dumps(data, default=str)))

        log.info("[%s] Saved (in:%d out:%d cur:%d)", zone.upper(), data["totalin"], data["totalout"], data["totalcur"])
    except asyncio.TimeoutError:
        log.warning("[%s] Timeout", zone.upper())
    except Exception:
        log.exception("[%s] Error saat fetch_store", zone.upper())

async def zone_loop(cfg: dict):
    name = cfg["name"]
    in_devices = [d.strip() for d in os.getenv(cfg["in_env"], "").split(",") if d.strip()]
    out_devices = [d.strip() for d in os.getenv(cfg["out_env"], "").split(",") if d.strip()]
    interval = int(os.getenv(cfg["interval_env"], "30"))

    if not in_devices or not out_devices:
        log.warning("[%s] IN/OUT devices kosong", name.upper())
        return

    log.info("[%s] Interval: %ds", name.upper(), interval)
    while True:
        await fetch_and_store(name, in_devices, out_devices)
        await asyncio.sleep(interval)

async def run_worker():
    log.info("Worker start")
    await asyncio.gather(*(zone_loop(z) for z in ZONES))

def setup_graceful_shutdown(loop: asyncio.AbstractEventLoop):
    async def shutdown():
        log.info("Shutdown...")
        tasks = [t for t in asyncio.all_tasks(loop) if t is not asyncio.current_task()]
        for t in tasks:
            t.cancel()
        await asyncio.gather(*tasks, return_exceptions=True)
        loop.stop()

    for sig in (signal.SIGINT, signal.SIGTERM):
        try:
            loop.add_signal_handler(sig, lambda: asyncio.create_task(shutdown()))
        except NotImplementedError:
            pass

if __name__ == "__main__":
    if platform.system() == "Windows":
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    setup_graceful_shutdown(loop)

    try:
        loop.run_until_complete(run_worker())
    except (KeyboardInterrupt, asyncio.CancelledError):
        log.info("Exit gracefully")
    finally:
        loop.close()
