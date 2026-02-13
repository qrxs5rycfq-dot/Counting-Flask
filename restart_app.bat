@echo off
setlocal

echo [RESTART] Menghentikan aplikasi...
call stop_app.bat

echo.
echo [RESTART] Menunggu 2 detik sebelum memulai ulang...
timeout /t 2 /nobreak >nul

echo.
echo [RESTART] Menjalankan aplikasi kembali...
call start_app.bat

echo.
echo [RESTART] Selesai. Menutup dalam 5 detik...
timeout /t 5 /nobreak >nul
exit /b