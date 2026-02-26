# KB023 Flask 서버 실행 (가상환경 생성 + 의존성 설치 + 서버 기동)
# PowerShell에서: .\run.ps1
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$pythonExe = $null
foreach ($name in @('python', 'py')) {
    $c = Get-Command $name -ErrorAction SilentlyContinue
    if ($c) { $pythonExe = $c.Source; break }
}
if (-not $pythonExe) {
    $paths = @(
        "$env:LocalAppData\Programs\Python\Python314\python.exe",
        "$env:LocalAppData\Programs\Python\Python313\python.exe",
        "$env:LocalAppData\Programs\Python\Python312\python.exe",
        "${env:ProgramFiles}\Python314\python.exe",
        "${env:ProgramFiles}\Python313\python.exe"
    )
    foreach ($p in $paths) {
        if (Test-Path $p) { $pythonExe = $p; break }
    }
}
if (-not $pythonExe) {
    Write-Host "`n[오류] Python을 찾을 수 없습니다.`n" -ForegroundColor Red
    Write-Host "1. Python 설치: https://www.python.org/downloads/ (설치 시 Add Python to PATH 체크)"
    Write-Host "2. Windows 설정 > 앱 > 앱 실행 별칭 에서 python.exe 끄기"
    Write-Host "3. 터미널을 새로 연 뒤 .\run.ps1 다시 실행`n"
    exit 1
}

if (-not (Test-Path "venv")) {
    Write-Host "가상환경 생성 중..." -ForegroundColor Cyan
    & $pythonExe -m venv venv
    if ($LASTEXITCODE -ne 0) { Write-Host "가상환경 생성 실패." -ForegroundColor Red; exit 1 }
}
& .\venv\Scripts\Activate.ps1
Write-Host "의존성 설치 중..." -ForegroundColor Cyan
pip install -r requirements.txt -q
Write-Host "`n서버 시작 (종료: Ctrl+C)..." -ForegroundColor Green
python app.py
