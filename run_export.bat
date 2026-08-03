@echo off
cd /d "%~dp0"

echo ============================================================
echo  Excel to LaTeX Exporter
echo ============================================================

:: Cek apakah venv sudah ada, kalau belum buat dulu
if not exist ".venv\Scripts\python.exe" (
    echo [1/4] Membuat virtual environment...
    python -m venv .venv
) else (
    echo [1/4] Virtual environment sudah ada.
)

:: Aktifkan venv
echo [2/4] Mengaktifkan venv...
call .venv\Scripts\activate.bat

:: Install dependensi
echo [3/4] Menginstall openpyxl dan pandas...
pip install openpyxl pandas -q

:: Jalankan script
echo [4/4] Menjalankan excel_to_latex.py...
echo.
python excel_to_latex.py

echo.
echo ============================================================
pause