@echo off
REM PowerShell에서는 .\run.bat 로 실행하세요.
chcp 65001 >nul
cd /d "%~dp0"

set "PYEXE="
where py >nul 2>nul && set "PYEXE=py"
if not defined PYEXE where python >nul 2>nul && set "PYEXE=python"
if not defined PYEXE (
    for %%V in (314 313 312 311 310) do (
        if exist "%LocalAppData%\Programs\Python\Python%%V\python.exe" set "PYEXE=%LocalAppData%\Programs\Python\Python%%V\python.exe"
        if defined PYEXE goto :py_found
    )
    for %%V in (314 313 312 311 310) do (
        if exist "%ProgramFiles%\Python%%V\python.exe" set "PYEXE=%ProgramFiles%\Python%%V\python.exe"
        if defined PYEXE goto :py_found
    )
)
:py_found
if not defined PYEXE (
    echo.
    echo [오류] Python을 찾을 수 없습니다.
    echo.
    echo 1. Python 설치: https://www.python.org/downloads/
    echo    - 설치 시 "Add Python to PATH" 체크
    echo.
    echo 2. 이미 설치했다면: Windows 설정 ^> 앱 ^> 고급 앱 설정 ^> 앱 실행 별칭
    echo    - "python.exe", "python3.exe" 끄기 (Microsoft Store로 열리는 것 방지)
    echo.
    echo 3. 터미널(명령 프롬프트)을 새로 연 뒤 다시 실행하세요.
    echo.
    pause
    exit /b 1
)

if not exist "venv" (
    echo 가상환경 생성 중...
    %PYEXE% -m venv venv
    if errorlevel 1 (
        echo 가상환경 생성 실패.
        pause
        exit /b 1
    )
)
call venv\Scripts\activate.bat
if not exist "venv\Scripts\python.exe" (
    echo venv\Scripts\python.exe 없음. venv 폴더를 삭제한 뒤 다시 실행하세요.
    pause
    exit /b 1
)
echo 의존성 설치 중...
pip install -r requirements.txt -q
echo.
echo 서버 시작 (종료: Ctrl+C)...
python app.py
pause
