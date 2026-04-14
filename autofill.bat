@echo off
chcp 65001 >nul 2>&1
title 입찰 정량평가 자동입력 시스템
echo.
echo ========================================
echo   입찰 정량평가 자동입력 시스템
echo   (주)스페이스락
echo ========================================
echo.
python "%~dp0autofill.py"
