<#
.SYNOPSIS
    Перемикання цільових URL у файлі manifest.xml між localhost та GitHub Pages.
.EXAMPLE
    .\switch_mode.ps1 -Mode Local
.EXAMPLE
    .\switch_mode.ps1 -Mode GitHubPages -GitHubUrl "https://oromis777-dev.github.io/excel_addin"
#>
param(
    [ValidateSet("Local", "GitHubPages")]
    [string]$Mode = "Local",
    
    [string]$GitHubUrl = "https://oromis777-dev.github.io/excel_addin"
)

$manifestPath = Join-Path $PSScriptRoot "manifest.xml"
if (-not (Test-Path $manifestPath)) {
    Write-Error "Файл manifest.xml не знайдено за шляхом $manifestPath"
    exit 1
}

$content = Get-Content $manifestPath -Raw -Encoding UTF8

if ($Mode -eq "Local") {
    $targetBase = "https://localhost:3000"
    Write-Host "Перемикання маніфесту на локальний режим ($targetBase)..." -ForegroundColor Green
    
    # Заміна GitHub Pages URL на localhost
    $pattern = 'https://[a-zA-Z0-9\-\._]+github\.io/[a-zA-Z0-9\-\._]+'
    $content = $content -replace $pattern, $targetBase
} else {
    $targetBase = $GitHubUrl.TrimEnd('/')
    Write-Host "Перемикання маніфесту на GitHub Pages ($targetBase)..." -ForegroundColor Cyan
    
    # Заміна localhost на GitHub Pages
    $content = $content -replace 'https://localhost:3000', $targetBase
    $content = $content -replace 'http://localhost:3000', $targetBase
}

Set-Content -Path $manifestPath -Value $content -Encoding UTF8
Write-Host "✅ manifest.xml успішно оновлено для режиму '$Mode'!" -ForegroundColor Green
