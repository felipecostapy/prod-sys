@echo off
echo ================================
echo  Compilando ViaTop...
echo ================================
echo.

call venv\Scripts\activate

pip install pyinstaller

pyinstaller Sistema_Ordens.spec

echo.
echo ================================
echo  Executavel gerado em: dist\Sistema_Ordens\
echo ================================
pause
