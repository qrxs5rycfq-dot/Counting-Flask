import os, sys, asyncio, socket, logging, signal, threading
from typing import Final

from flask import Flask
from waitress import serve

from app.routes.main_routes import register_routes
from worker.tracker_worker import run_worker
from models.models import create_tables
from app.utils.single_instance import ensure_single_instance

class AppServer:
    def __init__(self, base_dir: str):
        self.base_dir = base_dir
        self.port = int(os.getenv("APP_PORT", 12345))
        self.secret = os.getenv("SECRET_KEY", "supersecret")

        # ── Logging ───────────────────────────────────────────
        os.makedirs(os.path.join(self.base_dir, "logs"), exist_ok=True)
        logging.basicConfig(
            filename=os.path.join(self.base_dir, "logs", "app.log"),
            level=logging.INFO,
            format="%(asctime)s [%(levelname)s] %(message)s",
        )
        self.log = logging.getLogger("AppServer")

        self._validate_env()
        create_tables()

        # ── Flask ────────────────────────────────────────────
        self.app = Flask(
            __name__,
            template_folder=os.path.join(self.base_dir, "templates"),
            static_folder=os.path.join(self.base_dir, "static"),
        )
        self.app.secret_key = self.secret
        self._prepare_upload_folder()
        register_routes(self.app)

    def _prepare_upload_folder(self) -> None:
        upl = os.path.join(self.app.static_folder, "uploads")
        os.makedirs(upl, exist_ok=True)
        self.app.config["UPLOAD_FOLDER"] = upl

    def _validate_env(self) -> None:
        required = [
            "API_URL", "URL_DEPT", "URL_ADD_PERSON", "ACCESS_TOKEN",
            "DATABASE_URL", "TITLE_HIJAU", "TITLE_MERAH"
        ]
        missing = [k for k in required if not os.getenv(k)]
        if missing:
            self.log.critical(f"[ENV] variabel hilang: {', '.join(missing)}")
            sys.exit(1)

    def _start_worker(self) -> None:
        threading.Thread(
            target=lambda: asyncio.run(run_worker()),
            name="TrackerWorker",
            daemon=True
        ).start()
        self.log.info("[Worker] tracker aktif")

    @staticmethod
    def _setup_signals() -> None:
        def clean(*_):
            if os.path.exists("app.pid"):
                os.remove("app.pid")
            sys.exit(0)

        signal.signal(signal.SIGINT, clean)
        signal.signal(signal.SIGTERM, clean)

    def run(self) -> None:
        ensure_single_instance(self.port, self.log)
        self._start_worker()
        self._setup_signals()

        with open("app.pid", "w") as f:
            f.write(str(os.getpid()))
        self.log.info("Server siap di http://localhost:%s", self.port)

        serve(self.app, host="0.0.0.0", port=self.port)
