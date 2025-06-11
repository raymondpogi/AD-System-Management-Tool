@echo off
:menu
cls
echo ===============================
echo WMIC Power Tool (Non-Admin)
echo ===============================
echo 1. System Information
echo 2. Serial Number
echo 3. Install/Uninstall Software
echo 4. Update Software
echo 5. Network Configuration
echo 6. Running Processes
echo 7. Installed Printers
echo 8. Add Printer
echo 9. Windows Updates
echo 10. Power Plan Settings
echo 11. Exit
echo ===============================
color 0A
set /p choice=Enter your choice (1-11): 
color 07
if %choice%==1 goto sysinfo
if %choice%==2 goto serial
if %choice%==3 goto software
if %choice%==4 goto updatesoftware
if %choice%==5 goto network
if %choice%==6 goto processes
if %choice%==7 goto printers
if %choice%==8 goto addprinter
if %choice%==9 goto updates
if %choice%==10 goto powerplan
if %choice%==11 exit

:sysinfo
cls
echo === System Information ===
echo --- Computer System ---
wmic computersystem get name, username, manufacturer, model, totalphysicalmemory /format:list
if %errorlevel% neq 0 (
    echo Error: Failed to retrieve computer system information. WMIC may be unavailable.
)
echo.
echo --- Operating System ---
wmic os get caption, version, buildnumber, lastbootuptime /format:list
if %errorlevel% neq 0 (
    echo Error: Failed to retrieve OS information. WMIC may be unavailable.
)
echo.
echo --- BIOS ---
wmic bios get manufacturer, version, releasedate /format:list
if %errorlevel% neq 0 (
    echo Error: Failed to retrieve BIOS information. WMIC may be unavailable.
)
echo.
echo --- Disks ---
wmic logicaldisk get deviceid, description, freespace, size /format:list
if %errorlevel% neq 0 (
    echo Error: Failed to retrieve disk information. WMIC may be unavailable.
)
pause
goto menu

:serial
cls
echo === Serial Number ===
wmic bios get serialnumber /format:list
if %errorlevel% neq 0 (
    echo Error: Failed to retrieve serial number. WMIC may be unavailable.
)
pause
goto menu

:software
cls
echo === Install/Uninstall Software (Sorted by Install Date, Oldest to Newest) ===
wmic product get name, version, vendor, installdate /format:csv | sort /+2 > temp.csv
type temp.csv
del temp.csv
if %errorlevel% neq 0 (
    echo Error: Failed to retrieve software list. WMIC may be unavailable or you lack permissions.
    pause
    goto menu
)
echo.
color 0A
set /p softwarename=Enter the software name to uninstall (or press Enter to skip): 
color 07
if "%softwarename%"=="" goto menu
echo Attempting to uninstall "%softwarename%"...
wmic product where name="%softwarename%" call uninstall
if %errorlevel%==0 (
    echo Software uninstalled successfully.
) else (
    echo Failed to uninstall software. This action may require admin privileges.
)
pause
goto menu

:updatesoftware
cls
echo === Available Software Updates (via winget) ===
winget upgrade
if %errorlevel% neq 0 (
    echo Error: Failed to retrieve software updates. winget may be unavailable or not installed.
    pause
    goto menu
)
echo.
color 0A
set /p package=Enter the package ID or name to update (or press Enter to skip): 
color 07
if "%package%"=="" goto menu
echo Attempting to update "%package%"...
winget upgrade --id "%package%"
if %errorlevel%==0 (
    echo Software updated successfully.
) else (
    echo Failed to update software. This action may require admin privileges or the package ID is invalid.
)
pause
goto menu

:network
cls
echo === Network Configuration ===
wmic nicconfig get caption, ipaddress, macaddress, dhcpenabled /format:list
if %errorlevel% neq 0 (
    echo Error: Failed to retrieve network configuration. WMIC may be unavailable.
)
pause
goto menu

:processes
cls
echo === Running Processes (Sorted by Name, A-Z) ===
wmic process get name, processid, executablepath /format:csv | sort /+2 > temp.csv
type temp.csv
del temp.csv
if %errorlevel% neq 0 (
    echo Error: Failed to retrieve process list. WMIC may be unavailable.
    pause
    goto menu
)
echo.
color 0A
set /p pid=Enter the Process ID to kill (or press Enter to skip): 
color 07
if "%pid%"=="" goto menu
echo Attempting to kill process ID %pid%...
taskkill /PID %pid% /F
if %errorlevel%==0 (
    echo Process terminated successfully.
) else (
    echo Failed to kill process. You may lack permissions or the process ID is invalid.
)
pause
goto menu

:printers
cls
echo === Installed Printers ===
wmic printer get name /format:list
if %errorlevel% neq 0 (
    echo Error: Failed to retrieve printer list. WMIC may be unavailable.
    pause
    goto menu
)
echo.
color 0A
set /p printername=Enter the printer name to delete (or press Enter to skip): 
color 07
if "%printername%"=="" goto menu
echo Attempting to delete printer "%printername%"...
rundll32 printui.dll,PrintUIEntry /dn /n "%printername%"
if %errorlevel%==0 (
    echo Printer deleted successfully.
) else (
    echo Failed to delete printer. This action may require admin privileges.
)
pause
goto menu

:addprinter
cls
echo === Add Printer ===
color 0A
set /p servername=Enter the print server name (e.g., \\servername): 
color 07
if "%servername%"=="" (
    echo Error: No server name entered.
    pause
    goto menu
)
echo.
echo === Shared Printers on %servername% ===
net view %servername%
if %errorlevel% neq 0 (
    echo Error: Failed to retrieve shared printers. Check server name or network connectivity.
    pause
    goto menu
)
echo.
color 0A
set /p printershare=Enter the shared printer name to add (e.g., Printer1): 
color 07
if "%printershare%"=="" (
    echo Error: No printer name entered.
    pause
    goto menu
)
echo Attempting to add printer %servername%\%printershare%...
rundll32 printui.dll,PrintUIEntry /in /n "%servername%\%printershare%"
if %errorlevel%==0 (
    echo Printer added successfully.
) else (
    echo Failed to add printer. This action may require admin privileges or the printer/server name is invalid.
)
pause
goto menu

:updates
cls
echo === Installed Windows Updates (Sorted by Install Date, Oldest to Newest) ===
wmic qfe get hotfixid, description, installedon /format:csv | sort /+4 > temp.csv
type temp.csv
del temp.csv
if %errorlevel% neq 0 (
    echo Error: Failed to retrieve Windows updates. WMIC may be unavailable.
    pause
    goto menu
)
echo.
color 0A
set /p kbnumber=Enter the KB number to uninstall (e.g., KB1234567, or press Enter to skip): 
color 07
if "%kbnumber%"=="" goto menu
echo Attempting to uninstall update %kbnumber%...
wusa /uninstall /kb:%kbnumber% /quiet /norestart
if %errorlevel%==0 (
    echo Update uninstalled successfully.
) else (
    echo Failed to uninstall update. This action may require admin privileges or the KB number is invalid.
)
pause
goto menu

:powerplan
cls
echo === Current Power Plan Settings ===
echo --- Active Power Plan ---
powercfg /l
if %errorlevel% neq 0 (
    echo Error: Failed to retrieve power plan list. powercfg may be unavailable.
)
echo.
echo --- Detailed Settings (Active Plan) ---
powercfg /q
if %errorlevel% neq 0 (
    echo Error: Failed to retrieve detailed power settings. powercfg may be unavailable.
)
echo.
color 0A
set /p optimize=Optimize power settings (set sleep/display/HDD to 0, power button/lid to Do Nothing)? (y/n): 
color 07
if /i "%optimize%"=="y" (
    echo Attempting to optimize power settings...
    powercfg /x -standby-timeout-ac 0
    powercfg /x -standby-timeout-dc 0
    powercfg /x -monitor-timeout-ac 0
    powercfg /x -monitor-timeout-dc 0
    powercfg /x -disk-timeout-ac 0
    powercfg /x -disk-timeout-dc 0
    powercfg /x -powerbuttonaction 0
    powercfg /x -lidaction 0
    if %errorlevel%==0 (
        echo Power settings optimized successfully.
    ) else (
        echo Failed to optimize power settings. This action may require admin privileges.
    )
) else (
    echo Optimization skipped.
)
pause
goto menu