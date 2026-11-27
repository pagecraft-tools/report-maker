<#
  보고서 제작기 — 새 버전 내보내기 (개발용)
      .\release.ps1 -Note "고친 내용"
  version.txt 를 오늘 날짜 버전으로 올리고 커밋·push 한다.
  push 되면 설치본들이 다음에 켤 때 자동으로 새 버전을 받는다.
#>
param([string]$Note = '', [switch]$NoPush)
# git 은 줄바꿈 경고 같은 안내도 stderr 로 내보낸다. PowerShell 5.1 은 이것을 오류로 감싸서
# 'Stop' 이면 거기서 스크립트가 죽는다(2026-11-27 실제로 커밋 전에 멈춤). 그래서 git 결과는
# 오류 개체가 아니라 종료 코드($LASTEXITCODE)로만 판단한다.
$ErrorActionPreference = 'Continue'
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $root

function Invoke-Git {
    # 출력은 글자로만 보여 주고(경고는 회색), 성공 여부는 종료 코드로 돌려준다
    & git @args 2>&1 | ForEach-Object {
        $line = "$_"
        if ($line -match 'LF will be replaced|CRLF will be replaced') { return }
        if ($_ -is [System.Management.Automation.ErrorRecord]) { Write-Host "  $line" -ForegroundColor DarkGray }
        else { Write-Host "  $line" }
    }
    return ($LASTEXITCODE -eq 0)
}

function Fail($msg) {
    Write-Host "  [오류] $msg" -ForegroundColor Red
    exit 1
}

$remote = (& git remote get-url origin 2>$null)
if ($remote -notmatch 'pagecraft-tools/report-maker') {
    Fail "origin 이 pagecraft-tools/report-maker 가 아닙니다: $remote"
}
# 커밋 작성자는 메일 숨김 주소로(개인 메일이 공개 저장소에 찍히지 않게)
$email = (& git config user.email 2>$null)
if ($email -notmatch 'users\.noreply\.github\.com$') {
    Fail "커밋 메일이 숨김 주소가 아닙니다($email). 이 폴더에서 git config user.email 을 noreply 주소로 정해 주세요."
}

$today = Get-Date -Format 'yyyy.MM.dd'
$old = ''
if (Test-Path version.txt) { $old = (Get-Content version.txt -Raw).Trim() }
$n = 1
if ($old -match ('^' + [regex]::Escape($today) + '-(\d+)$')) { $n = [int]$Matches[1] + 1 }
$ver = "$today-$n"
[IO.File]::WriteAllText((Join-Path $root 'version.txt'), $ver + "`n", (New-Object Text.UTF8Encoding($false)))
Write-Host "  버전 $old → $ver" -ForegroundColor Cyan

if (-not (Invoke-Git add -A)) { Fail 'git add 실패' }
$msg = if ($Note) { "v$ver $Note" } else { "v$ver" }
if (-not (Invoke-Git commit -q -m $msg)) { Fail '커밋 실패 (바뀐 파일이 없거나 git 오류)' }
if ($NoPush) {
    Write-Host '  (push 하지 않았습니다)' -ForegroundColor DarkGray
    exit 0
}
if (-not (Invoke-Git push origin main)) {
    Write-Host '  커밋까지는 끝났지만 GitHub에 올리지 못했습니다(로그인 정보 없음 등).' -ForegroundColor Yellow
    Write-Host '  터미널에서 git push origin main 을 실행하면 브라우저 로그인 뒤 올라갑니다.' -ForegroundColor Yellow
    exit 1
}
Write-Host '  끝. 3~5분 뒤부터 설치본이 켤 때 새 버전을 받습니다.' -ForegroundColor Green
