param(
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'
$project = Split-Path -Parent $PSScriptRoot
$source = Join-Path $project 'solution\source'
if (-not $OutputPath) {
    $OutputPath = Join-Path $project 'solution\BOMAgent_1_0_0_3_unmanaged.zip'
}

if (-not (Get-Command pac -ErrorAction SilentlyContinue)) {
    throw 'Install the Power Platform CLI (PAC) before packing the solution.'
}

& pac solution pack --folder $source --zipfile $OutputPath --packagetype Unmanaged
if ($LASTEXITCODE -ne 0) {
    throw 'PAC solution packaging failed.'
}
