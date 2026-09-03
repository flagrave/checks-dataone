@echo off
setlocal
rem Corre el check y regenera el tablero. Doble clic desde la carpeta.
rem "%~dp0" = la carpeta donde vive este .bat, asi anda aunque la muevas.
cd /d "%~dp0"

rem ----------------------------------------------------------------
rem Carpeta compartida donde publicamos el tablero para los owners de
rem cada negocio. Cambia con cada ciclo de cierre: editar SOLO esta
rem linea cuando pasemos a la carpeta del ciclo siguiente.
rem ----------------------------------------------------------------
set "DEST=C:\Users\GMERP\Bayer\Finance BP CS Conosur - Documents\CY26\Closing 2026.08\FC3 2026\Submission\Checks"

echo.
echo ==========================================================
echo   Check SBP 2027  -  query SAP BW vs hoja de trabajo
echo ==========================================================
echo.
echo Acordate de haber GUARDADO Checksibf.xlsx despues del refresh.
echo.

py check_sbp.py
if errorlevel 1 goto error
echo.

py dashboard.py
if errorlevel 1 goto error
echo.

rem --- publicar en la carpeta compartida ---
rem Se copia recien aca: si alguno de los dos scripts fallo, ya salimos por
rem :error y los owners siguen viendo la ultima version buena.
if not exist "%DEST%\" goto nodest

copy /y "Tablero_SBP2027.html" "%DEST%\Tablero_SBP2027.html" >nul
if errorlevel 1 goto nocopy

echo   Publicado en la carpeta compartida:
echo   %DEST%
echo.
goto abrir

:nodest
echo   *** NO se publico: la carpeta compartida no esta disponible.
echo       %DEST%
echo       Revisa que OneDrive este sincronizado, o que la carpeta del
echo       ciclo no haya cambiado (se edita al inicio de este .bat).
echo.
goto abrir

:nocopy
echo   *** NO se publico: fallo la copia. Puede estar abierto el archivo
echo       en la carpeta compartida, o faltar permisos de escritura.
echo.

:abrir
echo ----------------------------------------------------------
echo   Listo. Abriendo el tablero...
echo ----------------------------------------------------------
start "" "Tablero_SBP2027.html"
echo.
pause
exit /b 0

:error
echo.
echo **********************************************************
echo   FALLO. Mira el mensaje de arriba. Lo mas comun:
echo    - Checksibf.xlsx abierto con cambios sin guardar
echo    - cambio el orden de las columnas de la hoja "query"
echo    - se renombro alguna hoja del libro
echo **********************************************************
echo.
pause
exit /b 1
