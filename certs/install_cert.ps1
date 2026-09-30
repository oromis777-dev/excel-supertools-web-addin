# install_cert.ps1 - Installation of localhost SSL certificate for current user
$ErrorActionPreference = "Stop"

$cerPath = Join-Path $PSScriptRoot "localhost.cer"
if (-not (Test-Path $cerPath)) {
    $cerPath = "c:\Users\Morion\Projects\excel-supertools-web-addin - 1\certs\localhost.cer"
}

if (-not (Test-Path $cerPath)) {
    Write-Error "Certificate file not found: $cerPath"
    exit 1
}

Write-Host "Installing certificate into CurrentUser Root store..." -ForegroundColor Cyan
& C:\Windows\System32\certutil.exe -user -addstore Root $cerPath

Write-Host "`nCertificate installation completed!" -ForegroundColor Green
Write-Host "Now return to Excel and click 'Перезапустити' (Restart) in the taskpane." -ForegroundColor Yellow
