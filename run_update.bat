@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"
set "PY=C:\Users\Faye\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
set "GIT_SSH_COMMAND=C:/Users/Faye/.workbuddy/vendor/PortableGit/usr/bin/ssh.exe"

echo ============================================
echo   退费渠道数据看板 - 一键更新
echo ============================================
echo.
echo [1/3] 读取 Excel 并重新生成看板页面...
"%PY%" "%~dp0build_data.py"
if errorlevel 1 (
  echo [错误] 生成失败，请确认 Excel 已关闭、数据正常。
  pause
  exit /b 1
)

echo.
echo [2/3] 提交改动到本地仓库...
git add -A
git diff --cached --quiet
if not errorlevel 1 (
  echo 数据无变化，无需提交与推送。
  goto :done
)
git commit -m "更新退费渠道数据"

echo.
echo [3/3] 推送到 GitHub Pages...
git push
if errorlevel 1 (
  echo 首次推送失败（可能是网络/DNS 临时问题），重试中...
  git push
)
if errorlevel 1 (
  echo [错误] 推送仍失败，本地页面已生成，请稍后手动推送或重试。
  pause
  exit /b 1
)

:done
echo.
echo 完成！稍等 1-2 分钟，访问以下地址即可看到更新：
echo   https://faye0366.github.io/refund-dashboard/
echo   https://faye0366.github.io/refund-dashboard/compare.html
pause
