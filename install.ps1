<#
  보고서 제작기 — 설치 (설치.bat 이 이 파일을 받아 실행한다)

  1) 파이썬 확인 (없으면 설치 안내)
  2) 프로그램을 GitHub 에서 받아 %LOCALAPPDATA%\보고서제작기 에 둔다
  3) 필요한 파이썬 패키지 설치
  4) 바탕화면에 「보고서 제작기」 바로가기
  다시 실행해도 안전하다(최신으로 덮어씀, 내 설정은 유지).
#>
param([string]$Dir = '', [switch]$NoShortcut)   # -NoShortcut: 시험 설치용

$Repo = 'pagecraft-tools/report-maker'
$Branch = 'main'
$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
function Say($m, $c = 'Gray') { Write-Host $m -ForegroundColor $c }

Say ''
Say '  보고서 제작기 설치' 'Cyan'
Say '  ─────────────────────────────────────────' 'DarkGray'

# 1) 파이썬
function Find-Python {
    foreach ($n in @('Python313', 'Python312', 'Python311')) {
        $p = Join-Path $env:LOCALAPPDATA "Programs\Python\$n\python.exe"
        if (Test-Path -LiteralPath $p) { return $p }
    }
    $c = Get-Command python -ErrorAction SilentlyContinue
    if ($c -and $c.Source -notlike '*\WindowsApps\*') { return $c.Source }
    return $null
}
$py = Find-Python
if (-not $py) {
    Say '  [1/4] 파이썬이 없어 설치합니다 (몇 분 걸립니다)…' 'White'
    $w = Get-Command winget -ErrorAction SilentlyContinue
    if ($w) {
        & winget install -e --id Python.Python.3.12 --scope user --accept-package-agreements --accept-source-agreements
        $py = Find-Python
    }
    if (-not $py) {
        Say '  파이썬을 자동으로 설치하지 못했습니다.' 'Red'
        Say '  https://www.python.org/downloads/ 에서 설치한 뒤(설치 첫 화면의 "Add python.exe to PATH" 체크) 다시 실행하세요.' 'Gray'
        exit 1
    }
}
Say "  [1/4] 파이썬 확인: $py" 'Green'

# 2) 프로그램 받기
if (-not $Dir) { $Dir = Join-Path $env:LOCALAPPDATA '보고서제작기' }
Say "  [2/4] 프로그램 받는 중 → $Dir" 'White'
$tmp = Join-Path $env:TEMP ("report-maker-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
$zip = Join-Path $tmp 'src.zip'
Invoke-WebRequest -Uri "https://github.com/$Repo/archive/refs/heads/$Branch.zip" -OutFile $zip -UseBasicParsing
Expand-Archive -LiteralPath $zip -DestinationPath $tmp -Force
$src = Get-ChildItem -LiteralPath $tmp -Directory | Select-Object -First 1
New-Item -ItemType Directory -Path $Dir -Force | Out-Null
Get-ChildItem -LiteralPath $src.FullName -Force | Where-Object { $_.Name -notin @('settings.json', '.git') -and $_.Name -match '^[ -~]+$' } |
    ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $Dir -Recurse -Force }
Set-Content -LiteralPath (Join-Path $Dir 'installed.txt') -Value $Repo -Encoding ASCII
Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue

# 3) 패키지
Say '  [3/4] 필요한 구성요소 설치 중…' 'White'
$ErrorActionPreference = 'Continue'
& $py -m pip install --quiet --disable-pip-version-check -r (Join-Path $Dir 'requirements.txt')
$ErrorActionPreference = 'Stop'

# 4) 바로가기
if (-not $NoShortcut) {
$desk = [Environment]::GetFolderPath('Desktop')
$lnk = Join-Path $desk '보고서 제작기.lnk'
$sh = New-Object -ComObject WScript.Shell
$s = $sh.CreateShortcut($lnk)
$s.TargetPath = 'wscript.exe'
$s.Arguments = '"' + (Join-Path $Dir 'start.vbs') + '"'
$s.WorkingDirectory = $Dir
$s.IconLocation = (Join-Path $Dir 'icon.ico')
$s.Save()
Say '  [4/4] 바탕화면에 「보고서 제작기」 아이콘을 만들었습니다.' 'Green'
}

Say ''
Say '  설치 끝! 바탕화면의 「보고서 제작기」로 켜세요. 켤 때마다 자동으로 최신 버전이 됩니다.' 'Green'
Say ''
exit 0
