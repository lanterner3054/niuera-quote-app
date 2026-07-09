@echo off
cd /d "%~dp0"
title NIUERA Quote Workbench - running (do not close)

echo ============================================
echo    NIUERA Quote Workbench
echo ============================================
echo.
echo   This PC :  http://127.0.0.1:8000
echo   LAN     :  http://^<your-lan-ip^>:8000
echo.
echo   - Same-WiFi phones/PCs open the LAN address.
echo   - First run: allow Python through firewall (Private).
echo   - Cannot connect? quit Clash / turn off TUN, retry.
echo   - Close this window to stop.
echo ============================================
echo.

start "" "http://127.0.0.1:8000"
python -m uvicorn app:app --host 0.0.0.0 --port 8000

echo.
echo Service stopped. Press any key to close...
pause >nul
