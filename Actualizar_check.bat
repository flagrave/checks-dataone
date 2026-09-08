@echo off
setlocal
rem Corre el check y regenera el tablero. Doble clic desde la carpeta.
rem "%~dp0" = la carpeta donde vive este .bat, asi anda aunque la muevas.
cd /d "%~dp0"

rem ----------------------------------------------------------------
rem Lo unico que se edita al cambiar de ciclo: estas cinco lineas.
rem CICLO tiene que coincidir con una clave de CYCLES en check_sbp.py
rem (hoy: sbp / fc3); LIBRO y TABLERO son los nombres que usa ese ciclo.
rem DEST es la carpeta compartida donde publicamos para los owners.
rem ----------------------------------------------------------------
set "CICLO=fc3"
set "TITULO=Check FC3 2026  -  actual (ene-ago) y to go (sep-dic)"
set "LIBRO=ChecksibfFC3.xlsx"
set "TABLERO=Tablero_FC3_2026.html"
set "DEST=C:\Users\GMERP\Bayer\Finance BP CS Conosur - Documents\CY26\Closing 2026.08\FC3 2026\Submission\Checks"

echo.
echo ==========================================================
echo   %TITULO%
echo ==========================================================
echo.
echo Acordate de haber GUARDADO %LIBRO% despues del refresh.
echo La hoja "query" tiene que traer el periodo contable (Posting
echo period) y la hoja de trabajo, los doce meses cargados: de ahi
echo sale la separacion entre el actual y el to go.
echo.

py check_sbp.py --cycle %CICLO%
if errorlevel 1 goto error
echo.

py dashboard.py --cycle %CICLO%
if errorlevel 1 goto error
echo.

rem --- publicar en la carpeta compartida ---
rem Se copia recien aca: si alguno de los dos scripts fallo, ya salimos por
rem :error y los owners siguen viendo la ultima version buena.
if not exist "%DEST%\" goto nodest

copy /y "%TABLERO%" "%DEST%\%TABLERO%" >nul
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
start "" "%TABLERO%"
echo.
pause
exit /b 0

:error
echo.
echo **********************************************************
echo   FALLO. Mira el mensaje de arriba. Lo mas comun:
echo    - %LIBRO% abierto con cambios sin guardar
echo    - la hoja "query" sin la columna del periodo contable
echo    - cambiaron los titulos de las columnas de la hoja "query"
echo    - se renombro alguna hoja del libro
echo **********************************************************
echo.
pause
exit /b 1
