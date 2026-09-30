# fix_catalog.ps1 - Автоматичне виправлення мережевого шляху каталогу Excel
$ErrorActionPreference = "Stop"

$correctUncPath = "\\MORION_MD\Users\Morion\Projects\excel-supertools-web-addin - 1"

Write-Host "1. Перевірка доступності мережевого шляху: $correctUncPath" -ForegroundColor Cyan
if (-not (Test-Path $correctUncPath)) {
    # Fallback to localhost if computer name differs
    $correctUncPath = "\\localhost\Users\Morion\Projects\excel-supertools-web-addin - 1"
    Write-Host "   Спроба через localhost: $correctUncPath" -ForegroundColor Yellow
}

if (-not (Test-Path $correctUncPath)) {
    Write-Error "Мережевий шлях $correctUncPath недоступний! Перевірте спільний доступ."
    exit 1
}

$manifestOnShare = Join-Path $correctUncPath "manifest.xml"
if (-not (Test-Path $manifestOnShare)) {
    Write-Error "Файл manifest.xml не знайдено за мережевим шляхом: $manifestOnShare"
    exit 1
}
Write-Host "   Знайдено manifest.xml за адресою: $manifestOnShare" -ForegroundColor Green

# Оновлення реєстру HKCU
$trustedCatalogsPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
if (-not (Test-Path $trustedCatalogsPath)) {
    New-Item -Path $trustedCatalogsPath -Force | Out-Null
}

# Отримуємо існуючі каталоги або створюємо новий
$catalogs = Get-ChildItem $trustedCatalogsPath -ErrorAction SilentlyContinue

if ($catalogs.Count -gt 0) {
    foreach ($cat in $catalogs) {
        Write-Host "2. Виправляємо існуючий каталог: $($cat.PSChildName)" -ForegroundColor Cyan
        Set-ItemProperty -Path $cat.PSPath -Name "Url" -Value $correctUncPath -Type String -Force
        Set-ItemProperty -Path $cat.PSPath -Name "Flags" -Value 1 -Type DWord -Force
        Write-Host "   Встановлено Url = $correctUncPath" -ForegroundColor Green
        Write-Host "   Встановлено Flags = 1 (Показувати в меню)" -ForegroundColor Green
    }
} else {
    $guid = [guid]::NewGuid().ToString("B").ToUpper()
    $newCatPath = Join-Path $trustedCatalogsPath $guid
    New-Item -Path $newCatPath -Force | Out-Null
    Set-ItemProperty -Path $newCatPath -Name "Id" -Value $guid -Type String -Force
    Set-ItemProperty -Path $newCatPath -Name "Url" -Value $correctUncPath -Type String -Force
    Set-ItemProperty -Path $newCatPath -Name "Flags" -Value 1 -Type DWord -Force
    Write-Host "2. Створено новий каталог $guid із правильним шляхом." -ForegroundColor Green
}

# Очищення кешу Office WEF
Write-Host "`n3. Очищення закешованих метаданих Office WEF..." -ForegroundColor Cyan
$wefPath = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef"
if (Test-Path $wefPath) {
    try {
        Get-ChildItem -Path $wefPath -Recurse -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "   Кеш WEF успішно очищено!" -ForegroundColor Green
    } catch {
        Write-Host "   Деякі файли кешу зайняті процесом Excel. Закрийте Excel перед повторним запуском." -ForegroundColor Yellow
    }
}

Write-Host "`n==================================================" -ForegroundColor Magenta
Write-Host " ГОТОВО! НАЛАШТУВАННЯ УСПІШНО ВИПРАВЛЕНО" -ForegroundColor Green
Write-Host "==================================================" -ForegroundColor Magenta
Write-Host "Правильний каталог: $correctUncPath"
Write-Host "`nІнструкція:"
Write-Host "1. Закрийте повністю всі відкриті вікна Excel (якщо вони відкриті)."
Write-Host "2. Запустіть сервер: npm start (якщо він ще не запущений)."
Write-Host "3. Відкрийте Excel наново."
Write-Host "4. Перейдіть: Вставка -> Надбудови -> Мої надбудови -> СПІЛЬНА ПАПКА."
Write-Host "5. Натисніть кнопку 'Оновити' (у правому верхньому куті вікна надбудов)."
Write-Host "   Тепер там з'явиться 'ExcelSuperTools (Web Add-in)'!"
