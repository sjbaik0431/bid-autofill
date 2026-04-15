@echo off
chcp 949 >nul 2>&1
title 입찰 정량평가 자동입력 대시보드
cd /d "%~dp0"

rem 진단 로그 생성
set "LOGFILE=%~dp0실행로그.txt"
echo === 실행 시작: %DATE% %TIME% === > "%LOGFILE%"
echo 현재폴더: %CD% >> "%LOGFILE%"

cls
echo.
echo  ==================================================
echo    입찰 정량평가 자동입력 대시보드
echo    http://localhost:5000
echo  ==================================================
echo.
echo  * 진단 로그는 실행로그.txt 파일에 저장됩니다.
echo.

rem --- 1. Python 확인 (python 또는 py) ---
echo [1/4] Python 확인 중...
set "PYCMD="
where python >nul 2>&1 && set "PYCMD=python"
if not defined PYCMD (
    where py >nul 2>&1 && set "PYCMD=py"
)
if not defined PYCMD (
    echo  [오류] Python이 설치되어 있지 않거나 PATH에 등록되지 않았습니다.
    echo.
    echo  ----- 해결 방법 -----
    echo   1. https://www.python.org/downloads/ 접속
    echo   2. Python 3.10 이상 다운로드 및 설치
    echo   3. 설치 시 "Add Python to PATH" 옵션을 반드시 체크!
    echo.
    echo ERROR: Python not found in PATH >> "%LOGFILE%"
    echo.
    echo ---- 창이 닫히지 않습니다. 확인 후 닫으세요 ----
    cmd /k
)
echo   -^> %PYCMD% 사용 가능
echo PYTHON: %PYCMD% >> "%LOGFILE%"
%PYCMD% --version >> "%LOGFILE%" 2>&1

rem --- 2. 한컴오피스 확인 ---
echo [2/4] 한컴오피스 확인 중...
reg query "HKEY_CLASSES_ROOT\HWPFrame.HwpObject" >nul 2>&1
if errorlevel 1 (
    echo   -^> 한컴오피스 미설치 감지 ^(자동입력시 실제 작동은 안됨^)
    echo WARN: HWP not installed >> "%LOGFILE%"
) else (
    echo   -^> 한컴오피스 확인 완료
    echo OK: HWP installed >> "%LOGFILE%"
)

rem --- 3. 필수 패키지 확인 ---
echo [3/4] 필수 패키지 확인 중...
%PYCMD% -c "import flask, win32com, olefile, werkzeug" >nul 2>&1
if errorlevel 1 (
    echo   -^> 패키지 미설치. 자동 설치 시작 ^(1-2분 소요^)...
    echo INSTALLING packages... >> "%LOGFILE%"
    %PYCMD% -m pip install flask pywin32 olefile werkzeug >> "%LOGFILE%" 2>&1
    if errorlevel 1 (
        echo.
        echo  [오류] 패키지 설치 실패. 인터넷 연결 확인 필요.
        echo  실행로그.txt 파일을 확인하세요.
        echo.
        cmd /k
    )
    echo   -^> 패키지 설치 완료
    echo OK: packages installed >> "%LOGFILE%"
) else (
    echo   -^> 패키지 확인 완료
    echo OK: packages already installed >> "%LOGFILE%"
)

rem --- 4. 포트 5000 확인 ---
echo [4/4] 포트 확인 중...
netstat -ano | findstr ":5000 " | findstr "LISTENING" >nul 2>&1
if not errorlevel 1 (
    echo   -^> 이미 실행 중. 브라우저 재오픈.
    start "" "http://localhost:5000"
    timeout /t 3 >nul
    exit /b 0
)
echo   -^> 포트 5000 사용 가능

echo.
echo  --------------------------------------------------
echo    대시보드 시작. 브라우저가 자동으로 열립니다.
echo    종료: 이 창에서 Ctrl+C 를 누르세요.
echo  --------------------------------------------------
echo.
echo DASHBOARD STARTING... >> "%LOGFILE%"

rem --- 대시보드 실행 ---
%PYCMD% "%~dp0dashboard.py"

echo.
echo DASHBOARD EXITED with code %ERRORLEVEL% >> "%LOGFILE%"
if errorlevel 1 (
    echo.
    echo  [오류] 대시보드가 비정상 종료되었습니다.
    echo  실행로그.txt 파일에서 원인을 확인하세요.
    echo.
    cmd /k
)
