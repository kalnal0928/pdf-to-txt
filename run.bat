@echo off
echo ================================
echo     PDF to TXT 변환기 실행
echo ================================
echo.
echo 1. GUI 버전 실행
echo 2. 명령행 버전 도움말 보기
echo 3. 종료
echo.
set /p choice="선택하세요 (1-3): "

if "%choice%"=="1" (
    echo GUI 버전을 실행합니다...
    python pdf_to_txt_gui.py
) else if "%choice%"=="2" (
    echo.
    python pdf_to_txt.py
) else if "%choice%"=="3" (
    echo 종료합니다.
    exit /b
) else (
    echo 잘못된 선택입니다.
    pause
)

pause
