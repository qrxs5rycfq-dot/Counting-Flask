import socket
import threading
import webbrowser
import sys
import logging

CONTROL_PORT = 59781

def ensure_single_instance(port: int, logger: logging.Logger = None) -> None:
    """
    Pastikan hanya satu instance berjalan.
    - Instance pertama membuka browser dan menjalankan listener.
    - Instance kedua hanya mengirim sinyal untuk membuka browser di instance utama.
    """
    def open_browser():
        try:
            url = f"http://localhost:{port}"
            webbrowser.open_new_tab(url)
            if logger:
                logger.info(f"[Browser] Membuka: {url}")
        except Exception as e:
            if logger:
                logger.warning(f"[Browser] Gagal membuka browser: {e}")

    def listener():
        while True:
            try:
                conn, _ = sock.accept()
                with conn:
                    open_browser()
            except Exception as e:
                if logger:
                    logger.warning(f"[SingleInstance] Listener error: {e}")

    try:
        # Coba bind ke CONTROL_PORT — jika berhasil, berarti ini instance pertama
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.bind(("localhost", CONTROL_PORT))
        sock.listen(1)

        if logger:
            logger.info("[SingleInstance] Instance pertama, listener aktif.")
        threading.Thread(target=listener, daemon=True).start()

        open_browser()

    except OSError:
        # Sudah ada instance lain — kirim sinyal ke listener
        try:
            with socket.create_connection(("localhost", CONTROL_PORT), timeout=1) as s:
                s.send(b"open")
            if logger:
                logger.info("[SingleInstance] Kirim sinyal buka browser ke instance utama.")
        except Exception as e:
            if logger:
                logger.warning(f"[SingleInstance] Gagal kirim sinyal ke instance utama: {e}")
        sys.exit(0)
