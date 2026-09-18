@echo off
cd /d "%~dp0"
"C:\Users\Faye\.workbuddy\binaries\python\envs\default\Scripts\python.exe" "%~dp0build_data.py"
"C:\Users\Faye\.workbuddy\binaries\python\envs\default\Scripts\python.exe" "%~dp0sync_to_dist.py"
echo.
echo Done: dashboard rebuilt and synced to dist.
echo To publish the share link, send this in chat: chongxinbushu
pause
