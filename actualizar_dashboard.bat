@echo off
REM Regenera data.json desde el Excel (solo lectura) y publica los cambios en GitHub Pages.
REM Por defecto lee Movimiento\ENE-JUL\ENE-JUL_MOVIMIENTO_SAPPRINECT.xlsx (RUTA_POR_DEFECTO en generar_datos.py).
REM Uso: doble clic, o:  actualizar_dashboard.bat "C:\ruta\a\otro_archivo.xlsx"
cd /d "%~dp0"

if "%~1"=="" (
  python generar_datos.py
) else (
  python generar_datos.py "%~1"
)
if errorlevel 1 (
  echo.
  echo ERROR al leer el Excel. No se publico nada.
  pause
  exit /b 1
)

git add data.json
git diff --cached --quiet
if not errorlevel 1 (
  echo.
  echo Sin cambios en los datos. Nada que publicar.
  pause
  exit /b 0
)

git commit -m "Actualizar datos del dashboard %date% %time%"
git push
echo.
echo Listo. El dashboard se actualiza en 1-2 minutos.
pause
