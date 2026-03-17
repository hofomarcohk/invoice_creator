@echo off
echo ============================================
echo  Invoice Generator - Build Script
echo ============================================
echo.

:: Install / update dependencies
echo [1/3] Installing dependencies...
pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: pip install failed.
    pause & exit /b 1
)

echo.
echo [2/3] Building executable with PyInstaller...
python -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "InvoiceGenerator" ^
    --add-data "Invoice_Sample.xlsx;." ^
    invoice_generator.py

if errorlevel 1 (
    echo ERROR: PyInstaller build failed.
    pause & exit /b 1
)

echo.
echo [3/3] Done!
echo The executable is at:  dist\InvoiceGenerator.exe
echo.
pause
