@echo off
cd /d "%~dp0"
echo Setting up DCDMD Dashboard...
python -m pip install -r requirements.txt
echo Starting the application...
python -m streamlit run app.py
pause