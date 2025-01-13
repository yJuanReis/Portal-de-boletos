@echo off
cd /d "%~dp0"
pip install -r ../Portal-de-boletos/requirements.txt
python "automacao-portal-de-boletos.py"
exit
