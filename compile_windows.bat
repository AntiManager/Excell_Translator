# Создаем compile_windows.bat
@'
@echo off
chcp 65001
echo ========================================
echo Excel Translator - Компиляция для Windows
echo ========================================
echo.

echo Проверка установленных пакетов...
pip list | findstr "pandas"
pip list | findstr "openpyxl" 
pip list | findstr "deep-translator"
pip list | findstr "requests"
pip list | findstr "tqdm"
pip list | findstr "pyinstaller"

echo.
echo Установка недостающих зависимостей...
pip install -r requirements.txt

echo.
echo Компиляция приложения...
pyinstaller --onefile --windowed ^
  --name=ExcelTranslator ^
  --hidden-import=deep_translator ^
  --hidden-import=deep_translator.google ^
  --hidden-import=deep_translator.base ^
  --hidden-import=pandas ^
  --hidden-import=openpyxl ^
  --hidden-import=requests ^
  --hidden-import=tqdm ^
  --clean ^
  main.py translator.py

if %errorlevel% == 0 (
    echo.
    echo ========================================
    echo ✓ КОМПИЛЯЦИЯ УСПЕШНО ЗАВЕРШЕНА!
    echo Исполняемый файл: dist\ExcelTranslator.exe
    echo ========================================
) else (
    echo.
    echo ✗ Ошибка компиляции
)

pause
'@ | Out-File -FilePath "compile_windows.bat" -Encoding utf8
