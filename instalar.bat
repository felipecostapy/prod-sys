@echo off
echo ================================
echo  Instalando ambiente ViaTop...
echo ================================
echo.

python -m venv venv
call venv\Scripts\activate

echo Instalando dependencias...
pip install -r requirements.txt

echo.
echo ================================
echo  Instalacao concluida!
echo  Use iniciar.bat para compilar.
echo ================================
pause
