@echo off
title Instalador de Gargam Stock
color 0A
echo ===================================================
echo   Instalador de Gargam Stock para Windows
echo ===================================================
echo.

:: Comprobar si Node.js esta instalado
node -v >nul 2>&1
if %errorlevel% neq 0 (
    color 0C
    echo [ERROR] Node.js no esta instalado en este equipo.
    echo.
    echo Para que la aplicacion funcione, necesitas instalar Node.js.
    echo 1. Ve a https://nodejs.org/
    echo 2. Descarga e instala la version "LTS"
    echo 3. Vuelve a ejecutar este archivo Setup.bat
    echo.
    pause
    exit /b
)

echo [1/3] Instalando dependencias del sistema...
cd ..
call npm install

echo.
echo [2/3] Compilando la aplicacion...
call npm run build

echo.
echo [3/3] Creando acceso directo en el Escritorio...
set LAUNCHER_DIR=%cd%\Windows_Setup
set SHORTCUT_SCRIPT=%LAUNCHER_DIR%\CreateShortcut.vbs
set TARGET_BAT=%LAUNCHER_DIR%\Iniciar_Gargam_Stock.bat

:: Crear el archivo .bat que iniciara la app
echo @echo off > "%TARGET_BAT%"
echo title Gargam Stock - Servidor >> "%TARGET_BAT%"
echo color 0B >> "%TARGET_BAT%"
echo echo =================================================== >> "%TARGET_BAT%"
echo echo   Gargam Stock esta en ejecucion >> "%TARGET_BAT%"
echo echo   NO CIERRES ESTA VENTANA MIENTRAS USES LA APP >> "%TARGET_BAT%"
echo echo =================================================== >> "%TARGET_BAT%"
echo echo. >> "%TARGET_BAT%"
echo cd "%%~dp0\.." >> "%TARGET_BAT%"
echo start http://localhost:3000 >> "%TARGET_BAT%"
echo call npm run dev >> "%TARGET_BAT%"

:: Crear el script VBS para el acceso directo
echo Set oWS = WScript.CreateObject("WScript.Shell") > "%SHORTCUT_SCRIPT%"
echo sLinkFile = oWS.SpecialFolders("Desktop") ^& "\Gargam Stock.lnk" >> "%SHORTCUT_SCRIPT%"
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> "%SHORTCUT_SCRIPT%"
echo oLink.TargetPath = "%TARGET_BAT%" >> "%SHORTCUT_SCRIPT%"
echo oLink.WorkingDirectory = "%LAUNCHER_DIR%" >> "%SHORTCUT_SCRIPT%"
echo oLink.Description = "Iniciar Gargam Stock" >> "%SHORTCUT_SCRIPT%"
echo oLink.Save >> "%SHORTCUT_SCRIPT%"

cscript /nologo "%SHORTCUT_SCRIPT%"
del "%SHORTCUT_SCRIPT%"

echo.
color 0A
echo ===================================================
echo   INSTALACION COMPLETADA CON EXITO
echo ===================================================
echo Se ha creado un acceso directo "Gargam Stock" en tu escritorio.
echo Haz doble clic en el para abrir la aplicacion.
echo.
pause
