@echo off
setlocal
cd /d "%~dp0"
set "ANACONDA_PYTHON=C:\Users\Omkar\anaconda3\python.exe"

echo Starting Social Analytics Pipeline app...
echo Local URL should open at http://localhost:8501
"%ANACONDA_PYTHON%" -m streamlit run "%~dp0app.py" --server.port 8501 --server.address localhost
pause
