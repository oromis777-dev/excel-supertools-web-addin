<#
.SYNOPSIS
    Видалення реєстрації Office Web Add-in із реєстру HKCU.
#>

$regKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer\ExcelSuperTools"

if (Test-Path $regKey) {
    Remove-Item -Path $regKey -Recurse -Force
    Write-Host "✅ Реєстрацію надбудови ExcelSuperTools видалено з реєстру ($regKey)." -ForegroundColor Green
    Write-Host "Перезапустіть Excel для оновлення стрічки." -ForegroundColor Yellow
} else {
    Write-Host "Гілка $regKey не знайдена в реєстрі." -ForegroundColor Cyan
}
