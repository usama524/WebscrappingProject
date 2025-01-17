@echo off
:: Run the app.exe invisibly using NirCmd
E:\Projects\Python Projects\goldline\vrm-expiry-webapp-new\tools\nircmd.exe exec hide dist\app.exe   :: Run the app.exe invisibly

:: Start the inactivity timer (Close app after 60 seconds)
set /a close_threshold=3600  :: 60 seconds before closing app.exe

:: Wait until Flask app is available (HTTP server responds)
:wait_for_flask
curl -s http://127.0.0.1:5000 > nul
if errorlevel 1 (
    timeout /t 1 > nul
    goto wait_for_flask
)

:: Once Flask is up, open the browser
start http://127.0.0.1:5000

:: Wait for 60 seconds before killing the app.exe
timeout /t %close_threshold% > nul

:: Kill the app.exe process after 60 seconds
echo Killing app.exe after %close_threshold% seconds...
for /f "tokens=2" %%a in ('tasklist /fi "imagename eq app.exe" /nh') do (
    echo Killing app.exe with PID %%a
    taskkill /f /pid %%a
)

:: Exit the script
::exit
pause