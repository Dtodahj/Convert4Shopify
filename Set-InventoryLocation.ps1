$ErrorActionPreference = 'Stop'
$ToolRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ConfigPath = Join-Path $ToolRoot 'mapping.json'

if (-not (Test-Path -LiteralPath $ConfigPath)) {
    throw 'mapping.json is missing beside the setup tool.'
}

Write-Host ''
Write-Host 'SHOPIFY INVENTORY LOCATION SETUP' -ForegroundColor Cyan
Write-Host 'In Shopify, open Settings > Locations and copy the exact active location name.'
Write-Host ''
$Location = Read-Host 'Enter the exact Shopify location name'
if ([string]::IsNullOrWhiteSpace($Location)) {
    throw 'No location was entered. Nothing was changed.'
}

$Config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
$Config.inventory.location = $Location.Trim()
$Json = $Config | ConvertTo-Json -Depth 10
$Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[IO.File]::WriteAllText($ConfigPath, $Json + [Environment]::NewLine, $Utf8NoBom)

Write-Host ''
Write-Host ("Saved inventory location: {0}" -f $Config.inventory.location) -ForegroundColor Green
Write-Host 'Setup is complete. You can now run Convert Invoices.cmd.'
