<#
.SYNOPSIS
  Installs the enterprise-slides Copilot agent and its helper library (per-user, no admin needed).
#>
$ErrorActionPreference = 'Stop'
$root   = Split-Path -Parent $MyInvocation.MyCommand.Path
$copilot = Join-Path $env:USERPROFILE '.copilot'
$assets = Join-Path $copilot 'slide-assets'

New-Item -ItemType Directory -Force (Join-Path $copilot 'agents') | Out-Null
foreach ($d in 'templates', 'doodles', 'fonts', 'lib') {
    New-Item -ItemType Directory -Force (Join-Path $assets $d) | Out-Null
}

Copy-Item (Join-Path $root 'agents\enterprise-slides.agent.md') (Join-Path $copilot 'agents') -Force
Copy-Item (Join-Path $root 'lib\slide_helpers.py') (Join-Path $assets 'lib') -Force

$styleFile = Join-Path $assets 'style.json'
if (-not (Test-Path $styleFile)) {
    Copy-Item (Join-Path $root 'style.example.json') $styleFile
}

Write-Host "Installed agent  -> $copilot\agents\enterprise-slides.agent.md"
Write-Host "Installed helpers -> $assets\lib\slide_helpers.py"
Write-Host "Drop your .pptx/.potx templates in $assets\templates, doodles in \doodles, fonts in \fonts."

$py = Get-Command python -ErrorAction SilentlyContinue
$real = $py -and ($py.Source -notlike '*WindowsApps*')
if (-not $real) {
    $local = Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\python.exe'
    if (Test-Path $local) { $py = $local; $real = $true }
}
if ($real) {
    $exe = if ($py -is [string]) { $py } else { $py.Source }
    & $exe -m pip install --quiet --disable-pip-version-check python-pptx lxml Pillow
    & $exe (Join-Path $root 'tests\smoke_test.py')
} else {
    Write-Warning "No real Python found. Run: winget install Python.Python.3.12 --scope user  (then re-run install.ps1)"
}
