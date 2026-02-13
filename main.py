from app.utils.path import get_base_dir
import os, sys
from dotenv import load_dotenv

BASE_DIR = get_base_dir()
sys.path.insert(0, BASE_DIR)

# .env dinamis
if getattr(sys, 'frozen', False):
    env_path = os.path.join(os.path.dirname(sys.executable), ".env")
else:
    env_path = os.path.join(BASE_DIR, ".env")

load_dotenv(env_path)

from app.core.server import AppServer

if __name__ == "__main__":
    AppServer(BASE_DIR).run()