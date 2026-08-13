@echo off
chcp 65001 >nul
echo ============================================
echo   Excel API Test - Update & Push to GitHub
echo ============================================
echo.

cd /d "%~dp0"

echo [1/4] git add
git add .
echo.

echo [2/4] git commit
git commit -m "update: 优化断言逻辑(嵌套字段查找) + Allure报告 + header配置化 + 日志滚动修复"
echo.

echo [3/4] git push
git push origin main
echo.

echo ============================================
echo   Done! 查看你的仓库:
echo   https://github.com/jalo-code/excelapitest
echo ============================================
pause
