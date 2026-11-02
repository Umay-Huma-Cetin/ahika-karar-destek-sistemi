@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo === AHIKA KDS v6.0 derleniyor ===
python -m pip install --upgrade pyinstaller pandas scikit-learn matplotlib reportlab openpyxl tkintermapview cryptography
python -m PyInstaller --noconfirm analiz.spec
if errorlevel 1 (
  echo DERLEME BASARISIZ. Yukaridaki hatayi kontrol edin.
  pause
  exit /b 1
)
if not exist yedekler mkdir yedekler
if exist "AHİKA KDS.exe" move /Y "AHİKA KDS.exe" "yedekler\AHİKA KDS_v5_onceki.exe" >nul
copy /Y "dist\analiz.exe" "AHİKA KDS.exe" >nul
echo.
echo Tamamlandi: "AHİKA KDS.exe" guncellendi.
pause
