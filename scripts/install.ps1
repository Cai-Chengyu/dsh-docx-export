# dsh-docx-export installer (Windows PowerShell 5.1+ / pwsh)
#
# Installs the Word-export plugin into a DSH profile through the official CLI
# and verifies that the bundle got registered:
#   dsh plugin --profile <profile> add <spec>
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File scripts/install.ps1
#   powershell -ExecutionPolicy Bypass -File scripts/install.ps1 -Profile desktop -Version 1.0.0 -Restart
#
# Parameters (all optional):
#   -Profile  target profile name (default: desktop)
#   -Version  npm version or range (default: latest) — ignored for -FromRepo / -FromPath
#   -FromRepo install straight from GitHub (github:<owner>/<repo>) instead of npm
#   -FromPath install from a local directory or .tgz
#   -Restart  try `pm2 restart dsh-web` after a successful install
#   -DryRun   print the commands without running them
#
# Environment:
#   DSH_HOME  default %USERPROFILE%\.dsh
#   DSH_CMD   override the dsh executable; defaults to PATH, then the
#             Windows desktop install, then `npx @deepseek-ai/dsh`
param(
  [string]$Profile = 'desktop',
  [string]$Version = 'latest',
  [string]$FromRepo = '',
  [string]$FromPath = '',
  [switch]$Restart,
  [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
$PKG = 'dsh-docx-export'

function Say([string]$m) { Write-Host "[install] $m" -ForegroundColor Green }
function Warn([string]$m) { Write-Host "[warn] $m" -ForegroundColor Yellow }
function Die([string]$m) { Write-Host "[error] $m" -ForegroundColor Red; exit 1 }

$DSH_HOME = if ($env:DSH_HOME) { $env:DSH_HOME } elseif ($env:USERPROFILE) { Join-Path $env:USERPROFILE '.dsh' } else { Join-Path $HOME '.dsh' }
$PROFILE_DIR = Join-Path $DSH_HOME "profiles\$Profile"

# --- resolve the CLI ---------------------------------------------------------
function Get-DshCli {
  if ($env:DSH_CMD) { return @{ Exe = $env:DSH_CMD; Args = @() } }
  if (Get-Command dsh -ErrorAction SilentlyContinue) { return @{ Exe = 'dsh'; Args = @() } }
  $desktop = Join-Path $env:LOCALAPPDATA 'Programs\DeepSeek Harness\resources\runtime\cli\bin\dsh.cmd'
  if (Test-Path $desktop) { return @{ Exe = $desktop; Args = @() } }
  if (Get-Command npx -ErrorAction SilentlyContinue) {
    return @{ Exe = 'npx'; Args = @('-y', '--package', '@deepseek-ai/dsh', 'dsh') }
  }
  return $null
}

# --- pick the install spec ---------------------------------------------------
if ($FromPath) {
  if (-not (Test-Path $FromPath)) { Die "path not found: $FromPath" }
  $spec = (Resolve-Path $FromPath).Path
} elseif ($FromRepo) {
  $spec = if ($FromRepo -match '^(github:|git\+|https?://)') { $FromRepo } else { "github:$FromRepo" }
} else {
  $spec = if ($Version -eq 'latest') { $PKG } else { "$PKG@$Version" }
}

$cli = Get-DshCli
if (-not $cli) { Die 'no dsh CLI found; install DSH, or set DSH_CMD / add npx to PATH.' }
if (-not (Test-Path (Join-Path $PROFILE_DIR 'package.json'))) {
  Warn "profile directory $PROFILE_DIR is not initialized yet; run DSH once with that profile first."
}

$cliArgs = @($cli.Args) + @('plugin', '--profile', $Profile, 'add', $spec)
Say ("about to run: {0} {1}" -f $cli.Exe, ($cliArgs -join ' '))
if ($DryRun) { Say '[dry-run] nothing was executed'; exit 0 }

& $cli.Exe @cliArgs
if ($LASTEXITCODE -ne 0) {
  Warn 'the install command failed. Common causes:'
  Warn '  - DSH Desktop is running and holds the profile during a fresh install; retry after it finishes or quit it.'
  Warn '  - the package name/version does not exist on npm yet (publish first, or use -FromPath / -FromRepo).'
  exit 1
}

# --- verify the bundle is selected ------------------------------------------
$manifest = Get-Content -Raw (Join-Path $PROFILE_DIR 'package.json') | ConvertFrom-Json
$bundleName = if ($FromPath) { (Get-Content -Raw (Join-Path $spec 'package.json') | ConvertFrom-Json).name } else { $PKG }
$bundles = @($manifest.dsh.profile.bundles)
if ($bundles -notcontains $bundleName) {
  Warn "$bundleName is not listed in dsh.profile.bundles — the mount is not registered."
  Warn "Add it by hand to $PROFILE_DIR\package.json, or re-run the install."
  exit 1
}
Say "bundle registered: $bundleName (mounts on the next DSH start)"

if ($Restart) {
  if (Get-Command pm2 -ErrorAction SilentlyContinue) {
    Say 'restarting dsh-web through pm2 ...'
    pm2 restart dsh-web
  } else {
    Warn 'pm2 not found; restart DSH by hand (desktop app: quit and reopen).'
  }
} else {
  Say 'next step: restart DSH (desktop app: quit and reopen), then ask for a Word export.'
}
