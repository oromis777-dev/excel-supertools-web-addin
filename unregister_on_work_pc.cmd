@echo off
chcp 65001 >nul
echo Видалення реєстрації ExcelSuperTools з реєстру користувача...
"%SystemRoot%\System32\reg.exe" delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer\ExcelSuperTools" /f >nul 2>&1
echo Готово. Надбудову вимкнено.
pause
