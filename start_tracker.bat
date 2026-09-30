@echo off
cd /d "%~dp0"
echo ========================================================
echo   SteamDB Cloudflare Bypass Wrapper
echo ========================================================
echo.
echo 1. Opening Microsoft Edge...
echo 2. PLEASE SOLVE THE CLOUDFLARE CHECKBOX IN THE BROWSER.
echo 3. Wait for the SteamDB table to fully load on the screen.
echo 4. Close steamdb_upcoming_tracker.xlsx and steamdb_upcoming_tracker_zh-CN.xlsx if they are open.
echo.
start "" msedge --remote-debugging-port=9222 --user-data-dir="%~dp0edge_profile" "https://steamdb.info/upcoming/?sort=followers_desc"
echo Once you see the table loaded with games...
pause
echo.
echo 5. Taking over the browser and extracting data...
node steamdb_tracker.js --remote
if errorlevel 1 (
    echo.
    echo Tracker failed. Review the error above and try again.
    pause
    exit /b 1
)
echo.
echo Extraction complete. Check steamdb_upcoming_tracker.xlsx and/or steamdb_upcoming_tracker_zh-CN.xlsx.
pause
