@echo off
setlocal
cd /d "%~dp0"

REM ====================================================
REM  WorkReport AutoFill launcher v1.3 (CP949)
REM  실행할 Streamlit 앱 파일명 (이 파일과 같은 폴더에 두세요)
set "MAIN_SCRIPT=autofill_app.py"
REM ====================================================

echo ====================================================
echo  WorkReport AutoFill 시작 준비 중...
echo ====================================================
echo.

if not exist "%MAIN_SCRIPT%" (
    echo [오류] %MAIN_SCRIPT% 파일을 찾을 수 없습니다.
    echo 이 배치 파일과 같은 폴더에 앱 파일이 있어야 합니다.
    pause
    exit /b 1
)

REM ---- 1. 실제 Python 찾기 (MS 스토어 가짜 별칭은 걸러냄) ----
set "PYTHON_CMD="

py -3 -c "exit()" >nul 2>&1
if not errorlevel 1 set "PYTHON_CMD=py -3"

if not defined PYTHON_CMD (
    python -c "exit()" >nul 2>&1
    if not errorlevel 1 set "PYTHON_CMD=python"
)

REM ---- 2. Python이 없으면 자동 설치 ----
if not defined PYTHON_CMD (
    echo [안내] Python이 설치되어 있지 않습니다. 자동 설치를 시작합니다.
    echo.
    where winget >nul 2>&1
    if not errorlevel 1 (
        winget install --id Python.Python.3.12 -e --accept-source-agreements --accept-package-agreements
    ) else (
        echo winget이 없어 python.org에서 설치 파일을 내려받습니다. 잠시 기다려 주세요...
        powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol='Tls12'; Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.10/python-3.12.10-amd64.exe' -OutFile ($env:TEMP + '\python-installer.exe')"
        if not exist "%TEMP%\python-installer.exe" (
            echo [오류] 설치 파일 다운로드에 실패했습니다. 인터넷 연결을 확인하세요.
            pause
            exit /b 1
        )
        echo Python을 설치하는 중입니다. 설치 창이 닫힐 때까지 기다려 주세요...
        "%TEMP%\python-installer.exe" /passive InstallAllUsers=0 PrependPath=1 Include_launcher=1
    )

    if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
        set "PYTHON_CMD="%LOCALAPPDATA%\Programs\Python\Python312\python.exe""
    )
)

if not defined PYTHON_CMD (
    echo.
    echo [안내] Python 설치가 끝났습니다. 이 창을 닫고 배치 파일을 다시 실행해 주세요.
    pause
    exit /b 1
)

REM ---- 3. 가상환경 생성 및 필요한 패키지 설치 ----
if not exist ".venv\Scripts\python.exe" (
    echo 가상환경을 생성하는 중...
    %PYTHON_CMD% -m venv .venv
)
if not exist ".venv\Scripts\python.exe" (
    echo [오류] 가상환경 생성에 실패했습니다.
    pause
    exit /b 1
)

if exist "requirements.txt" (
    echo 필요한 패키지를 확인/설치하는 중... (최초 실행 시 몇 분 걸립니다)
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt --disable-pip-version-check --quiet
    if errorlevel 1 (
        echo [오류] 패키지 설치에 실패했습니다. 인터넷 연결을 확인한 뒤 다시 실행해 주세요.
        pause
        exit /b 1
    )
)

REM ---- 4. 프로그램 실행 ----
echo.
echo WorkReport AutoFill을 실행합니다. 잠시 후 브라우저가 자동으로 열립니다...
echo (종료하려면 이 창을 닫으세요)
echo.
".venv\Scripts\python.exe" -m streamlit run "%MAIN_SCRIPT%"
set "EXIT_CODE=%errorlevel%"
echo.
if not "%EXIT_CODE%"=="0" (
    echo [오류] 프로그램이 오류 코드 %EXIT_CODE% 로 종료되었습니다.
)
pause
exit /b %EXIT_CODE%
