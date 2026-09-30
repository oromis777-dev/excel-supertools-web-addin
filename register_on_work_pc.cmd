@echo off
chcp 65001 >nul
echo ========================================================
echo   Реєстрація ExcelSuperTools на робочому комп'ютері
echo ========================================================
echo.

set "SCRIPT_DIR=%~dp0"
set "MANIFEST_PATH=%SCRIPT_DIR%manifest.xml"

if not exist "%MANIFEST_PATH%" (
    echo [ПОМИЛКА] Файл manifest.xml не знайдено поруч із цим скриптом!
    echo Переконайтеся, що файл manifest.xml лежить в одній папці з цим bat-файлом.
    echo.
    pause
    exit /b 1
)

echo Знайдено маніфест: %MANIFEST_PATH%
echo.
echo Запис у реєстр користувача (HKCU - не вимагає прав адміністратора)...

"%SystemRoot%\System32\reg.exe" add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer\ExcelSuperTools" /v ManifestPath /t REG_SZ /d "%MANIFEST_PATH%" /f >nul 2>&1

if %errorlevel% equ 0 (
    echo.
    echo [УСПІХ] Надбудову успішно зареєстровано!
    echo.
    echo Наступні кроки:
    echo 1. Повністю закрийте всі відкриті вікна Excel (якщо відкриті).
    echo 2. Запустіть Excel наново.
    echo 3. На стрічці Excel з'явиться вкладка "ExcelSuperTools".
    echo.
) else (
    echo.
    echo [ПОПЕРЕДЖЕННЯ] Не вдалося записати в реєстр через reg.exe (можливо, заборонено політикою).
    echo Скористайтеся інструкцією для завантаження через "Upload My Add-in" або через макрос VBA.
    echo.
)

pause
