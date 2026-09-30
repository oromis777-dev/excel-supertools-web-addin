<#
.SYNOPSIS
    Реєстрація Office Web Add-in у системному реєстрі HKCU без прав адміністратора.
.DESCRIPTION
    Скрипт записує шлях до manifest.xml у гілку розробника Office WEF Developer.
    Після виконання надбудова автоматично з'являється у стрічці Excel.
#>

$ErrorActionPreference = "Stop"

# Автоматичне визначення абсолютного шляху до manifest.xml у поточній папці
$manifestPath = Join-Path $PSScriptRoot "manifest.xml"

if (-not (Test-Path $manifestPath)) {
    Write-Error "Файл manifest.xml не знайдено за шляхом: $manifestPath"
    exit 1
}

$regKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer\ExcelSuperTools"

try {
    Write-Host "Реєстрація надбудови ExcelSuperTools у реєстрі користувача (HKCU)..." -ForegroundColor Cyan
    
    if (-not (Test-Path $regKey)) {
        New-Item -Path $regKey -Force | Out-Null
    }

    Set-ItemProperty -Path $regKey -Name "ManifestPath" -Value $manifestPath -Type String -Force

    Write-Host "`n✅ Надбудову успішно зареєстровано в реєстрі!" -ForegroundColor Green
    Write-Host "   Гілка: $regKey"
    Write-Host "   Параметр ManifestPath: $manifestPath"
    Write-Host "`n⚠️  Щоб надбудова з'явилася на стрічці:" -ForegroundColor Yellow
    Write-Host "   1. Переконайтеся, що сервер запущено (npm start)."
    Write-Host "   2. Повністю закрийте всі вікна Excel (перевірте, щоб у Диспетчері завдань не лишилося excel.exe)."
    Write-Host "   3. Відкрийте Excel наново — на стрічці з'явиться вкладка 'ExcelSuperTools'."
} catch {
    Write-Error "Помилка запису в реєстр: $_"
}
