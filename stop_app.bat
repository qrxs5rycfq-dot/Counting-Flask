@echo off
setlocal

echo [STOP] Membaca PID dari app.pid...

REM === Cek apakah file PID ada ===
if not exist app.pid (
    echo [ERROR] File PID tidak ditemukan.
    timeout /t 5 >nul
    exit /b 1
)

REM === Baca PID dari file ===
set /p PID=<app.pid

REM === Validasi PID hanya angka ===
echo %PID% | findstr /r "^[0-9][0-9]*$" >nul
if errorlevel 1 (
    echo [ERROR] PID tidak valid: %PID%
    timeout /t 5 >nul
    exit /b 1
)

REM === Coba hentikan proses ===
echo [KILL] Menutup proses dengan PID: %PID%...
taskkill /PID %PID% /F >nul 2>&1

REM === Cek status ===
if %ERRORLEVEL%==0 (
    echo [DONE] Proses berhasil dihentikan.
    del app.pid
) else (
    echo [ERROR] Gagal menghentikan proses. Mungkin proses sudah tidak aktif.
)

echo Menutup dalam 5 detik...
timeout /t 5 >nul
exit /b