@echo off
cd /d "%~dp0"
set "PY=C:\Users\Faye\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
set "GIT_SSH_COMMAND=C:/Users/Faye/.workbuddy/vendor/PortableGit/usr/bin/ssh.exe"

echo ============================================
echo   Refund Dashboard - One-click Update
echo ============================================
echo.
echo [1/3] Rebuilding pages from Excel ...
"%PY%" "%~dp0build_data.py"
if errorlevel 1 (
  echo [ERROR] Build failed. Make sure Excel is closed and data is valid.
  pause
  exit /b 1
)

echo.
echo [2/3] Committing changes ...
git add -A
git diff --cached --quiet
if not errorlevel 1 (
  echo No data changes. Skip commit and push.
  goto :done
)
git commit -m "update refund data"

echo.
echo [3/3] Pushing to GitHub Pages ...
git push
if errorlevel 1 (
  echo Push failed (network/DNS?), retrying ...
  git push
)
if errorlevel 1 (
  echo [ERROR] Push still failed. Local pages are rebuilt, push manually later.
  pause
  exit /b 1
)

:done
echo.
echo Done! Wait 1-2 minutes, then visit:
echo   https://faye0366.github.io/refund-dashboard/
echo   https://faye0366.github.io/refund-dashboard/compare.html
pause
