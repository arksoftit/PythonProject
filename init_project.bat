@echo off
set /p project_name="Nombre del proyecto: "
mkdir "%project_name%"
cd "%project_name%"
python -m venv venv
echo Proyecto '%project_name%' creado con un entorno virtual.
pause