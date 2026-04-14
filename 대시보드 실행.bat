@echo off
chcp 65001 >nul 2>&1
title 입찰 정량평가 자동입력 대시보드
echo.
echo  ========================================
echo    입찰 정량평가 자동입력 대시보드
echo    브라우저가 자동으로 열립니다...
echo  ========================================
echo.
python "%~dp0dashboard.py"
pause
