@echo off
setlocal

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

pyinstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --name ExcelAutomation ^
  --collect-all openpyxl ^
  main.py

echo.
echo Build finalizado. Executavel em: dist\ExcelAutomation.exe
endlocal
