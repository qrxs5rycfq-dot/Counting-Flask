@echo off
setlocal

REM === Konfigurasi Path ===
set PY_PATH=%LocalAppData%\Programs\Python\Python313
set SCRIPT_PATH=C:\counting\main.py
set LOG_PATH=C:\counting\start_app.log

REM === Tulis log saat mulai ===
echo [%date% %time%] Starting app... >> "%LOG_PATH%"

REM === Cek apakah Pythonw ada ===
if not exist "%PY_PATH%\pythonw.exe" (
    echo [%date% %time%] ERROR: pythonw.exe tidak ditemukan di %PY_PATH% >> "%LOG_PATH%"
    echo Pythonw tidak ditemukan. Periksa PATH!
    timeout /t 5 >nul
    exit /b 1
)

REM === Cek apakah script ada ===
if not exist "%SCRIPT_PATH%" (
    echo [%date% %time%] ERROR: Script %SCRIPT_PATH% tidak ditemukan >> "%LOG_PATH%"
    echo Script tidak ditemukan.
    timeout /t 5 >nul
    exit /b 1
)

REM === Jalankan secara silent ===
start "" "%PY_PATH%\pythonw.exe" "%SCRIPT_PATH%"
if %errorlevel%==0 (
    echo [%date% %time%] App started silently. >> "%LOG_PATH%"
    echo App started successfully. Closing in 5 seconds...
) else (
    echo [%date% %time%] ERROR: Gagal menjalankan aplikasi >> "%LOG_PATH%"
    echo Gagal menjalankan aplikasi.
)

timeout /t 5 >nul
exit /b