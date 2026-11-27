<#
  보고서 제작기 — 자동 업데이트 (설치본에서만 돈다: installed.txt 가 있을 때)

  켤 때마다 GitHub 의 version.txt 를 보고, 바뀌었으면 저장소 zip 을 받아 프로그램 파일만 갈아 끼운다.
  원칙: 업데이트 때문에 프로그램이 안 켜지는 일은 없어야 한다. 무슨 실패든 조용히 넘어가고 지금 버전으로 켠다.
  내 설정(settings.json)·작업기록은 건드리지 않는다.
#>
param([string]$AppDir = $PSScriptRoot)

$Repo = 'pagecraft-tools/report-maker'
$Branch = 'main'

try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

function Get-Text([string]$uri) {
    $r = Invoke-WebRequest -Uri $uri -TimeoutSec 8 -UseBasicParsing
    return [Text.Encoding]::UTF8.GetString($r.RawContentStream.ToArray()).TrimStart([char]0xFEFF).Trim()
}

try {
    $local = ''
    $vf = Join-Path $AppDir 'version.txt'
    if (Test-Path -LiteralPath $vf) { $local = (Get-Content -LiteralPath $vf -Raw).Trim() }
    $stamp = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $remote = Get-Text "https://raw.githubusercontent.com/$Repo/$Branch/version.txt?t=$stamp"
    if (-not $remote -or $remote -eq $local) { exit 0 }

    $tmp = Join-Path $env:TEMP ("report-maker-" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tmp -Force | Out-Null
    $zip = Join-Path $tmp 'src.zip'
    Invoke-WebRequest -Uri "https://github.com/$Repo/archive/refs/heads/$Branch.zip" -OutFile $zip -TimeoutSec 60 -UseBasicParsing
    Expand-Archive -LiteralPath $zip -DestinationPath $tmp -Force
    $src = Get-ChildItem -LiteralPath $tmp -Directory | Select-Object -First 1
    if (-not $src -or -not (Test-Path -LiteralPath (Join-Path $src.FullName 'gui.py'))) { throw 'zip 내용이 이상함' }

    Get-ChildItem -LiteralPath $src.FullName -Force | Where-Object { $_.Name -notin @('settings.json', 'installed.txt', '.git') -and $_.Name -match '^[ -~]+$' } |
        ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $AppDir -Recurse -Force }

    # 새 버전에 필요한 파이썬 패키지가 늘었으면 깐다(조용히)
    $req = Join-Path $AppDir 'requirements.txt'
    if (Test-Path -LiteralPath $req) {
        $py = Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\python.exe'
        if (-not (Test-Path -LiteralPath $py)) { $py = Join-Path $env:LOCALAPPDATA 'Programs\Python\Python313\python.exe' }
        if (-not (Test-Path -LiteralPath $py)) { $py = 'python' }
        & $py -m pip install --quiet --disable-pip-version-check -r $req 2>$null | Out-Null
    }
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
} catch {
    try { Add-Content -LiteralPath (Join-Path $AppDir 'update.log') -Value ("{0}  {1}" -f (Get-Date), $_.Exception.Message) -Encoding UTF8 } catch {}
}
exit 0
