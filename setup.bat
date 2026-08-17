@echo off
echo A criar ambiente virtual...
py -3.12 -m venv venv
call venv\Scripts\activate
echo A instalar dependencias...
pip install -r requirements.txt
echo.
echo Setup concluido! Para correr o script:
echo   venv\Scripts\activate
echo   python duplicate_finder.py "C:\Caminho\A\Analisar" --move
pause