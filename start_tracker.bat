@echo off
echo ========================================================
echo   SteamDB Cloudflare Bypass Wrapper
echo ========================================================
echo.
echo 1. Opening Microsoft Edge...
echo 2. PLEASE SOLVE THE CLOUDFLARE CHECKBOX IN THE BROWSER.
echo 3. Wait for the SteamDB table to fully load on the screen.
echo.
start msedge --remote-debugging-port=9222 --user-data-dir="%~dp0\edge_profile" "https://steamdb.info/upcoming/?week=2026W40"
echo Once you see the table loaded with games...
pause
echo.
echo 4. Taking over the browser and extracting data...
node steamdb_tracker.js --remote
echo.
echo Extraction complete! You can now close the browser window.
pause
