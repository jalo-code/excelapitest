@echo off
chcp 65001 >nul
echo ============================================
echo   Excel API Test - Push to GitHub
echo ============================================
echo.

cd /d "%~dp0"

echo [1/6] git init
git init
echo.

echo [2/6] git add
git add .
echo.

echo [3/6] git commit
git commit -m "feat: Excel驱动的接口自动化测试框架（含Allure报告）"
echo.

echo [4/6] git branch -M main
git branch -M main
echo.

echo [5/6] git remote add origin
git remote remove origin 2>nul
git remote add origin https://github.com/jalo-code/excelapitest.git
echo.

echo [6/6] git push
git push -u origin main
echo.

echo ============================================
echo   Done! 查看你的仓库:
echo   https://github.com/jalo-code/excelapitest
echo ============================================
pause
