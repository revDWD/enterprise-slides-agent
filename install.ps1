<#
.SYNOPSIS
  Installs the enterprise-slides Copilot agent and helper library per-user.
.DESCRIPTION
  Python dependencies are isolated in the user's slide-assets virtual environment.
  The installer does not elevate privileges or change machine-wide policy.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$copilot = Join-Path $env:USERPROFILE '.copilot'
$assets = Join-Path $copilot 'slide-assets'
$venv = Join-Path $assets 'venv'
$venvPython = Join-Path $venv 'Scripts\python.exe'

New-Item -ItemType Directory -Force (Join-Path $copilot 'agents') | Out-Null
foreach ($directory in 'templates', 'doodles', 'fonts', 'lib') {
    New-Item -ItemType Directory -Force (Join-Path $assets $directory) | Out-Null
}

$styleFile = Join-Path $assets 'style.json'
if (-not (Test-Path $styleFile)) {
    Copy-Item (Join-Path $root 'style.example.json') $styleFile
}

$pythonCommand = Get-Command python -CommandType Application -ErrorAction SilentlyContinue
$python = $null
if ($pythonCommand -and $pythonCommand.Source -notlike '*WindowsApps*') {
    $python = $pythonCommand.Source
}
if (-not $python) {
    $localPython = Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\python.exe'
    if (Test-Path $localPython) {
        $python = $localPython
    }
}

if ($python) {
    & $python -c 'import sys; raise SystemExit(0 if sys.version_info >= (3, 12) else 1)'
    if ($LASTEXITCODE -ne 0) {
        throw 'Python 3.12 or newer is required. Install an approved per-user Python runtime, then rerun this installer.'
    }

    if (-not (Test-Path $venvPython)) {
        & $python -m venv $venv
        if ($LASTEXITCODE -ne 0) {
            throw "Could not create the isolated Python environment at $venv."
        }
    }

    & $venvPython -m pip install --disable-pip-version-check -r (Join-Path $root 'requirements.txt')
    if ($LASTEXITCODE -ne 0) {
        throw 'Dependency installation failed. Check your approved Python package source and enterprise policy; no workaround was attempted.'
    }

    & $venvPython (Join-Path $root 'tests\smoke_test.py')
    if ($LASTEXITCODE -ne 0) {
        throw 'The slide-helper smoke test failed. The agent files were not copied.'
    }
} else {
    Write-Warning 'No approved Python 3.12+ runtime was found. Install one through your approved software channel, then rerun this installer.'
    Write-Warning 'Python dependencies and the smoke test were not run.'
}

Copy-Item (Join-Path $root 'agents\enterprise-slides.agent.md') (Join-Path $copilot 'agents') -Force
Copy-Item (Join-Path $root 'lib\slide_helpers.py') (Join-Path $assets 'lib') -Force

Write-Host "Installed agent   -> $copilot\agents\enterprise-slides.agent.md"
Write-Host "Installed helpers -> $assets\lib\slide_helpers.py"
if (Test-Path $venvPython) {
    Write-Host "Isolated Python   -> $venvPython"
}
Write-Host "Drop approved .pptx/.potx templates in $assets\templates; restart Copilot to discover the agent."
