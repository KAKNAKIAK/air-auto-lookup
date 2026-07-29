<#
publish_release.ps1
  항공자동조회 배포 및 GitHub Release / latest.json 원스톱 스크립트.

  사용예:
    .\publish_release.ps1                 # 빌드부터 배포까지 한 번에 실행
    .\publish_release.ps1 -SkipBuild      # 기존 빌드파일로 배포만 수행
    .\publish_release.ps1 -Version v1.0.1 # 특정 버전으로 지정 배포
#>
param(
    [string]$Version = "",
    [switch]$SkipBuild,
    [string]$Repo = "KAKNAKIAK/air-auto-lookup"
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ReleaseRoot = Join-Path $ProjectRoot "release"
$RootLatest  = Join-Path $ProjectRoot "latest.json"
$Utf8NoBom = [System.Text.UTF8Encoding]::new($false)

# 1) 버전 자동 판별
if (-not $Version) {
    $bs = Get-Content -LiteralPath (Join-Path $ProjectRoot "build_setup.ps1") -Raw
    if ($bs -match '\$Version\s*=\s*"([^"]+)"') { $Version = $Matches[1] }
}
if (-not $Version) { throw "버전을 확인할 수 없습니다. -Version v1.0.x 로 지정하세요." }
Write-Host "[publish] 배포 대상 버전: $Version"

# 2) 셋업 빌드 실행
if (-not $SkipBuild) {
    Write-Host "[publish] build_setup.ps1 실행 중..."
    & (Join-Path $ProjectRoot "build_setup.ps1")
    if ($LASTEXITCODE -ne 0) { throw "빌드 실패" }
}

$Setup = Join-Path $ReleaseRoot ("AirAutoLookup_Setup_{0}.exe" -f $Version)
if (-not (Test-Path -LiteralPath $Setup)) {
    throw "설치 파일이 존재하지 않습니다: $Setup (build_setup.ps1 확인 필요)"
}

# 3) GitHub 릴리즈 생성 및 설치본 자산 업로드
$AssetName = Split-Path -Leaf $Setup
$ErrorActionPreference = "Continue"
gh release view $Version --repo $Repo *> $null
$releaseExists = ($LASTEXITCODE -eq 0)
$ErrorActionPreference = "Stop"

if ($releaseExists) {
    Write-Host "[publish] 기존 릴리즈 자산 업로드 (덮어쓰기)..."
    gh release upload $Version $Setup --repo $Repo --clobber
} else {
    Write-Host "[publish] 새 GitHub 릴리즈 생성 및 자산 업로드..."
    gh release create $Version $Setup --repo $Repo --title $Version --notes ("항공자동조회 {0} 릴리즈" -f $Version)
}
if ($LASTEXITCODE -ne 0) { throw "GitHub 릴리즈 업로드 실패" }

# 4) SHA256 실측 및 매니페스트 동기화
$DownloadUrl = "https://github.com/{0}/releases/download/{1}/{2}" -f $Repo, $Version, $AssetName
$AssetHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $Setup).Hash
Write-Host ("[publish] 셋업 파일 SHA256: {0}" -f $AssetHash)

$j = Get-Content -LiteralPath $RootLatest -Raw -Encoding UTF8 | ConvertFrom-Json
$j.version = $Version
$j.download_url = $DownloadUrl
$j.sha256 = $AssetHash

$RootLatestJson = $j | ConvertTo-Json -Depth 5
[System.IO.File]::WriteAllText($RootLatest, $RootLatestJson, $Utf8NoBom)
$ReleaseLatest = Join-Path $ReleaseRoot "latest.json"
[System.IO.File]::WriteAllText($ReleaseLatest, $RootLatestJson, $Utf8NoBom)
Write-Host "[publish] latest.json 매니페스트 동기화 완료"

# 5) 커밋 & 푸시 (Git 저장소인 경우)
Push-Location $ProjectRoot
try {
    if (Test-Path -LiteralPath (Join-Path $ProjectRoot ".git")) {
        git add latest.json
        git commit -m ("release: {0} 매니페스트 동기화" -f $Version)
        if ($LASTEXITCODE -ne 0) { Write-Host "[publish] 커밋할 변경이 없거나 커밋 완료됨" }
        git push origin HEAD
        if ($LASTEXITCODE -ne 0) { Write-Host "[publish] git push 경고 (원격 연결 확인 필요)" }
    } else {
        Write-Host "[publish] 참고: 현재 폴더가 git 저장소가 아닙니다. (.git 생성 필요시 git init 실행)"
    }
} finally {
    Pop-Location
}

Write-Host ""
Write-Host ("==========================================")
Write-Host ("배포 성공: {0}" -f $Version)
Write-Host ("  설치 파일: {0}" -f $Setup)
Write-Host ("  다운로드 : {0}" -f $DownloadUrl)
Write-Host ("  SHA256   : {0}" -f $AssetHash)
Write-Host ("==========================================")
