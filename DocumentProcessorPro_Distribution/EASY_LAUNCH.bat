@echo off
title Document Processor Pro - Easy Launcher
color 0A
echo.
echo ===============================================
echo   Document Processor Pro - Professional Edition
echo ===============================================
echo.
echo Starting your document processing application...
echo.

REM Check if the main executable exists
if not exist "DocumentProcessorPro.exe" (
    echo ERROR: DocumentProcessorPro.exe not found!
    echo.
    echo Please make sure you're running this from the correct folder.
    echo The folder should contain DocumentProcessorPro.exe
    echo.
    pause
    exit /b 1
)

REM Display system info
echo System: %OS% %PROCESSOR_ARCHITECTURE%
echo User: %USERNAME%
echo Time: %DATE% %TIME%
echo.

REM Try to run the application
echo Launching Document Processor Pro...
echo.
start "" "DocumentProcessorPro.exe"

REM Check if it started successfully
timeout /t 3 /nobreak >nul
echo Application should now be running with your custom icon!
echo.
echo If you see the application window, you can close this launcher.
echo If there are any issues, please check the troubleshooting guide.
echo.
echo Launcher will close in 10 seconds...
timeout /t 10 /nobreak >nul
