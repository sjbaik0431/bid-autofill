@echo off
chcp 949 >nul 2>&1
title 바탕화면 바로가기 생성

echo.
echo  ==================================================
echo    입찰 자동입력 대시보드 - 바로가기 생성
echo  ==================================================
echo.
echo  바탕화면에 "입찰 자동입력" 바로가기를 만듭니다.
echo.

set "TARGET=%~dp0대시보드 실행.bat"
set "SHORTCUT=%USERPROFILE%\Desktop\입찰 자동입력.lnk"
set "ICONPATH=%SystemRoot%\System32\shell32.dll,44"

rem PowerShell로 바로가기 생성
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$s=(New-Object -ComObject WScript.Shell).CreateShortcut('%SHORTCUT%'); ^
   $s.TargetPath='%TARGET%'; ^
   $s.WorkingDirectory='%~dp0'; ^
   $s.IconLocation='%ICONPATH%'; ^
   $s.WindowStyle=1; ^
   $s.Description='입찰 정량평가 자동입력 대시보드'; ^
   $s.Save()"

if errorlevel 1 (
    echo  [오류] 바로가기 생성 실패
    pause
    exit /b 1
)

echo  [완료] 바탕화면에 "입찰 자동입력" 아이콘이 생성되었습니다.
echo.
echo  이제 이 아이콘을 더블클릭하면 대시보드가 자동 실행됩니다.
echo.
pause
