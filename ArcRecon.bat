@echo off
mode con: cols=60 lines=20
color 0A
title Arc Recon Toolbox

set "OfficePath=C:\Program Files (x86)\ArcRecon"
set "ConfigFile=%OfficePath%\configuration.xml"
set "UninstallExe=%OfficePath%\Uninstall.exe"
echo.

:menu
cls
echo.
echo ==========================
echo     Arc Recon Toolbox
echo ==========================
echo.
echo 1. Activate Your OS
echo 2. Install Foffice
echo 3. Activate FOffice
echo 4. Remove OneDrive
echo 5. Run ChrisTitusTool (To Reduce Services)
echo 6. Activate Spicetify
echo 7. Exit the application
echo.

set /p choice=Enter your choice : 

if "%choice%"=="1" goto activate
if "%choice%"=="2" goto foffice
if "%choice%"=="3" goto office21
if "%choice%"=="4" goto onedrive
if "%choice%"=="5" goto ctt
if "%choice%"=="6" goto freespotify
if "%choice%"=="7" goto exiter

:activate
set /P achoice=Do you want to Activate Windows ? (Y/N): 
if /i "%achoice%" EQU "Y" goto windowsact
if /i "%achoice%" EQU "N" goto menu

:windowsact
    echo Select your Windows version:
    echo 1. Home    
    echo 2. Home Single Language
    echo 3. Education
    echo 4. Enterprise
    echo 5. Professional

    set /p wchoice=Enter your choice (1-5):

    if "%wchoice%"=="1" (
    set key=TX9XD-98N7V-6WMQ6-BX7FG-H8Q99
    goto activate
    )

    if "%wchoice%"=="2" (
    set key=W269N-WFGWX-YVC9B-4J6C9-T83GX
    goto activate
    )

    if "%wchoice%"=="3" (
    set key=NW6C2-QMPVW-D7KKK-3GKT6-VCFB2
    goto activate
    )

    if "%wchoice%"=="4" (
    set key=NPPR9-FWDCX-D2C8J-H872K-2YT43
    goto activate
    )

    if "%wchoice%"=="5" (
    set key=W269N-WFGWX-YVC9B-4J6C9-T83GX
    goto activate
    )

    :activate
    echo Installing license key...
    slmgr /skms kms8.msguides.com

    :retry
    slmgr /ipk %key%
    if errorlevel 0 (
        goto end
    )

    echo Error- Server busy. Retrying...
    timeout /t 5 >nul
    goto retry

    :end
    slmgr /ato
    echo Activation complete.
    Pause
    goto menu

:foffice
    set /p "Consent=Do you want to install Office 2021? (Y/N): "
    if /i "%Consent%" EQU "y" (
      echo Starting Office setup...
      echo.
    start /wait setup /configure "%ConfigFile%"

    echo Office setup completed.
    pause
    echo Exiting...
    goto menu
    ) else goto menu

:office21
      set /p "ActivateConsent=Do u want to activate the Office 2021 ? (Y/N): "
      if /i "%ActivateConsent%" EQU "y" (
      if exist "C:\Program Files\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office16"
      if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office16"
      for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x"
      cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
      cscript ospp.vbs /sethst:kms.msgang.com
      cscript ospp.vbs /act
      pause
      echo Exiting...
      pause
      goto menu
    ) else goto menu

:onedrive
    set /P onechoice=Do you want to remove Onedrive ? (Y/N): 
    if /i "%onechoice%"=="Y" (
        echo Uninstalling OneDrive...
        echo.

        REM Terminate any running OneDrive processes
        taskkill /f /im OneDrive.exe >nul 2>&1
        REM Uninstall OneDrive
        %SystemRoot%\SysWOW64\OneDriveSetup.exe /uninstall >nul 2>&1
        %SystemRoot%\System32\OneDriveSetup.exe /uninstall >nul 2>&1

        REM Remove OneDrive folders and registry keys
        rd "%UserProfile%\OneDrive" /Q /S >nul 2>&1
        rd "C:\OneDriveTemp" /Q /S >nul 2>&1
        rd "%LocalAppData%\Microsoft\OneDrive" /Q /S >nul 2>&1
        rd "%ProgramData%\Microsoft OneDrive" /Q /S >nul 2>&1

        reg delete "HKEY_CLASSES_ROOT\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}" /f >nul 2>&1
        reg delete "HKEY_CLASSES_ROOT\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}" /f >nul 2>&1
        reg delete "HKEY_CURRENT_USER\Software\Microsoft\OneDrive" /f >nul 2>&1
        reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}" /f >nul 2>&1
        reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Classes\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}" /f >nul 2>&1

        echo.
        echo OneDrive has been uninstalled.
        pause
        goto menu
    ) else goto menu


:ctt
    set /P cttchoice=Do you want to use the Chris Titus tool? (Y/N): 
    if /i "%cttchoice%"=="Y" (
        powershell -Command "iwr -useb https://christitus.com/win | iex" 
        pause
        goto menu
    ) else goto menu    

:freespotify
    echo Install Spotify and login in to your account before running this ...
    set /P spotchoice=Do you want to install Spicetify? (Y/N): 
    if /i "%spotchoice%"=="Y" (
        powershell -Command "iwr -useb https://raw.githubusercontent.com/spicetify/spicetify-cli/master/install.ps1 | iex" 
        pause
        powershell -Command "iwr -useb https://raw.githubusercontent.com/spicetify/spicetify-marketplace/main/resources/install.ps1 | iex"
        pause
        goto menu
    ) else goto menu

:exiter
    if exist "%OfficePath%"(
    start /wait "" "%UninstallExe%"
    echo Tool uninstallation completed.
    echo.
    echo Exiting...
    Exit
    ) else Exit
