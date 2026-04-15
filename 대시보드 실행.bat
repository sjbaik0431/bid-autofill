@echo off
chcp 65001 >nul 2>&1
title 입찰 정량평가 자동입력 대시보드

:: ================================================================
::  입찰 정량평가 자동입력 대시보드 실행기 v2
::  - Python 자동 확인
::  - 필수 패키지 자동 설치
::  - 브라우저 자동 열기
::  - 에러 발생 시 원인 안내
:: ================================================================

cd /d "%~dp0"

cls
echo.
echo  ╔══════════════════════════════════════════════════╗
echo  ║                                                  ║
echo  ║     입찰 정량평가 자동입력 대시보드              ║
echo  ║                                                  ║
echo  ║     http://localhost:5000                        ║
echo  ║                                                  ║
echo  ╚══════════════════════════════════════════════════╝
echo.

:: ── 1. Python 설치 확인 ──
where python >nul 2>&1
if errorlevel 1 (
    echo  [오류] Python이 설치되어 있지 않습니다.
    echo.
    echo  해결 방법:
    echo    1. https://www.python.org/downloads/ 접속
    echo    2. Python 3.10 이상 다운로드 및 설치
    echo    3. 설치 시 "Add Python to PATH" 반드시 체크!
    echo.
    pause
    exit /b 1
)

echo  [1/4] Python 확인 완료
echo.

:: ── 2. 한컴오피스 설치 확인 ──
reg query "HKEY_CLASSES_ROOT\HWPFrame.HwpObject" >nul 2>&1
if errorlevel 1 (
    echo  [경고] 한컴오피스가 설치되어 있지 않을 수 있습니다.
    echo         자동입력 기능은 한컴오피스가 있어야 동작합니다.
    echo.
    choice /C YN /M "  계속 진행하시겠습니까 (Y/N)"
    if errorlevel 2 exit /b 0
    echo.
) else (
    echo  [2/4] 한컴오피스 확인 완료
    echo.
)

:: ── 3. 필수 패키지 확인 ──
python -c "import flask, win32com, olefile" >nul 2>&1
if errorlevel 1 (
    echo  [3/4] 필수 패키지 설치 중... ^(최초 1회만^)
    echo.
    python -m pip install --quiet flask pywin32 olefile werkzeug
    if errorlevel 1 (
        echo.
        echo  [오류] 패키지 설치 실패. 인터넷 연결을 확인하세요.
        pause
        exit /b 1
    )
    echo  [3/4] 패키지 설치 완료
    echo.
) else (
    echo  [3/4] 필수 패키지 확인 완료
    echo.
)

:: ── 4. 포트 5000 사용 중인지 확인 ──
netstat -ano | findstr ":5000 " | findstr "LISTENING" >nul 2>&1
if not errorlevel 1 (
    echo  [알림] 포트 5000이 이미 사용 중입니다.
    echo         이미 대시보드가 실행 중일 수 있습니다.
    echo.
    start "" "http://localhost:5000"
    echo  브라우저에서 기존 창을 열었습니다.
    echo.
    timeout /t 3 >nul
    exit /b 0
)

echo  [4/4] 대시보드 시작 중...
echo.
echo  ──────────────────────────────────────────────────
echo    종료하려면 이 창에서 Ctrl+C 를 누르세요.
echo  ──────────────────────────────────────────────────
echo.

:: ── 대시보드 실행 ──
python "%~dp0dashboard.py"

if errorlevel 1 (
    echo.
    echo  [오류] 대시보드 실행 중 문제가 발생했습니다.
    echo.
    pause
)
