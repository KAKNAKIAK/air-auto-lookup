$ErrorActionPreference = "Stop"

$Version = "v1.0.8"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$SourceDir = Join-Path $ProjectRoot "source"
$ReleaseRoot = Join-Path $ProjectRoot "release"
$ReleaseDir = Join-Path $ReleaseRoot "AirAutoLookup"
$SetupOut = Join-Path $ReleaseRoot "AirAutoLookup_Setup_$Version.exe"
$SetupAlias = Join-Path $ReleaseRoot "AirAutoLookup_Setup.exe"
$TempInstallerDir = Join-Path ([System.IO.Path]::GetTempPath()) "AirAutoLookupInstallerTemp"
$Utf8NoBom = [System.Text.UTF8Encoding]::new($false)

function Reset-TempDir {
    param([Parameter(Mandatory = $true)][string]$Path)
    $TempRootFull = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath()).TrimEnd("\") + "\"
    $FullPath = [System.IO.Path]::GetFullPath($Path)
    if (-not $FullPath.StartsWith($TempRootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to modify a non-temp path: $FullPath"
    }
    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Recurse -Force
    }
    New-Item -ItemType Directory -Force -Path $Path | Out-Null
}

$PythonExe = "python.exe"

Push-Location $SourceDir
try {
    Write-Host "[1/4] PyInstaller 빌드 환경 확인 중..."
    & $PythonExe -m PyInstaller --version > $null 2> $null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "PyInstaller 설치 중..."
        & $PythonExe -m pip install pyinstaller
    }

    Write-Host "[2/4] 항공자동조회 실행 파일 빌드 중..."
    $TempPortableRoot = Join-Path ([System.IO.Path]::GetTempPath()) "AirAutoLookupPortableBuild"
    $TempPortableDist = Join-Path $TempPortableRoot "dist"
    $TempPortableBuild = Join-Path $TempPortableRoot "build"
    Reset-TempDir -Path $TempPortableRoot

    & $PythonExe -m PyInstaller `
        --noconfirm `
        --clean `
        --distpath $TempPortableDist `
        --workpath $TempPortableBuild `
        (Join-Path $SourceDir "항공자동조회.spec")

    if ($LASTEXITCODE -ne 0) {
        throw "항공자동조회 실행 파일 빌드 실패."
    }

    $ExeFile = Join-Path $TempPortableDist "항공자동조회.exe"
    if (-not (Test-Path -LiteralPath $ExeFile)) {
        throw "빌드된 실행 파일을 찾지 못했습니다: $ExeFile"
    }

    # 예시 샘플 4개 포함 hotels-manifest.json 보장
    $InitialManifest = @{
        schema = "air-auto-lookup-flight-masters-v1"
        updatedAt = (Get-Date -Format "yyyy-MM-ddTHH:mm:ss")
        flightMasters = @(
            @{ key="ICNGUM_LJ915_916"; rawKey="ICNGUM_LJ915/6"; origin="ICN"; destination="GUM"; route="ICN-GUM"; airline="LJ"; depFlight="LJ915"; retFlight="LJ916"; retDepartureTime=$null; defaultProductDays=5; fareRoute="ICN-GUM-LJ"; enabled=$true },
            @{ key="ICNGUM_LJ915_918"; rawKey="ICNGUM_LJ915/8"; origin="ICN"; destination="GUM"; route="ICN-GUM"; airline="LJ"; depFlight="LJ915"; retFlight="LJ918"; retDepartureTime=$null; defaultProductDays=5; fareRoute="ICN-GUM-LJ"; enabled=$true },
            @{ key="ICNGUM_LJ917_916"; rawKey="ICNGUM_LJ917/6"; origin="ICN"; destination="GUM"; route="ICN-GUM"; airline="LJ"; depFlight="LJ917"; retFlight="LJ917"; retDepartureTime=$null; defaultProductDays=5; fareRoute="ICN-GUM-LJ"; enabled=$true },
            @{ key="ICNGUM_LJ917_918"; rawKey="ICNGUM_LJ917/8"; origin="ICN"; destination="GUM"; route="ICN-GUM"; airline="LJ"; depFlight="LJ917"; retFlight="LJ918"; retDepartureTime=$null; defaultProductDays=5; fareRoute="ICN-GUM-LJ"; enabled=$true }
        )
    } | ConvertTo-Json -Depth 5
    [System.IO.File]::WriteAllText((Join-Path $TempPortableDist "hotels-manifest.json"), $InitialManifest, $Utf8NoBom)

    # release 디렉토리 갱신 및 동기화
    if (-not (Test-Path -LiteralPath $ReleaseRoot)) {
        New-Item -ItemType Directory -Force -Path $ReleaseRoot | Out-Null
    }
    New-Item -ItemType Directory -Force -Path $ReleaseDir | Out-Null
    try {
        Copy-Item -Path "$TempPortableDist\*" -Destination $ReleaseDir -Recurse -Force -ErrorAction Stop
    }
    catch {
        Write-Warning "실행 중인 앱 때문에 휴대용 release 폴더 갱신은 건너뜁니다. 셋업 파일 생성과 GitHub 배포는 계속 진행합니다."
    }

    Write-Host "[3/4] 셋업 페이로드 패키징 중..."
    Reset-TempDir -Path $TempInstallerDir
    
    $PayloadSrc = Join-Path $TempInstallerDir "payload"
    New-Item -ItemType Directory -Force -Path $PayloadSrc | Out-Null
    Copy-Item -Path "$TempPortableDist\*" -Destination $PayloadSrc -Recurse -Force -Exclude "hotels-manifest.json", "output", "logs"
    
    $PayloadZip = Join-Path $TempInstallerDir "payload.zip"
    Compress-Archive -Path "$PayloadSrc\*" -DestinationPath $PayloadZip -Force

    Write-Host "[4/4] 셋업 실행 파일 (Installer EXE) 생성 중..."
    if (Test-Path -LiteralPath $SetupOut) {
        Remove-Item -LiteralPath $SetupOut -Force
    }
    & $PythonExe -m PyInstaller `
        --noconfirm `
        --clean `
        --onefile `
        --windowed `
        --name "AirAutoLookup_Setup_$Version" `
        --icon (Join-Path $SourceDir "assets\air_auto_lookup_icon.ico") `
        --distpath $ReleaseRoot `
        --workpath (Join-Path $TempInstallerDir "build") `
        --specpath $TempInstallerDir `
        --add-data "$PayloadZip;." `
        (Join-Path $ProjectRoot "packaging\setup_installer.py")

    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $SetupOut)) {
        throw "셋업 실행 파일 빌드 실패."
    }

    Copy-Item -LiteralPath $SetupOut -Destination $SetupAlias -Force

    # latest.json 업데이트
    $LatestTemplate = Join-Path $ProjectRoot "latest.json"
    $SetupHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $SetupOut).Hash
    $DownloadUrl = "https://github.com/KAKNAKIAK/air-auto-lookup/releases/download/$Version/AirAutoLookup_Setup_$Version.exe"
    
    $LatestObj = @{
        version = $Version
        download_url = $DownloadUrl
        sha256 = $SetupHash
        release_notes = "항공자동조회 $Version 릴리즈 (빈 노선 매니페스트 기본 적용)"
    }
    $Json = $LatestObj | ConvertTo-Json -Depth 5
    $LatestOut = Join-Path $ReleaseRoot "latest.json"
    [System.IO.File]::WriteAllText($LatestOut, $Json, $Utf8NoBom)
    [System.IO.File]::WriteAllText($LatestTemplate, $Json, $Utf8NoBom)

    Write-Host "=========================================="
    Write-Host "빌드 완료!"
    Write-Host "설치 파일: $SetupOut"
    Write-Host "설치 별칭: $SetupAlias"
    Write-Host "매니페스트: $LatestTemplate"
    Write-Host "SHA256:    $SetupHash"
    Write-Host "=========================================="
}
finally {
    Pop-Location
}
