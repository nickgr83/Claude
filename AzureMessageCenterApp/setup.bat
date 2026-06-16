@echo off
echo === Azure Message Center Monitor - Setup ===
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Install Python 3.11+ from python.org
    pause
    exit /b 1
)

echo Installing dependencies...
python -m pip install -r requirements.txt

echo.
echo Done! Before running app.py, make sure to:
echo   1. Register an Azure AD app (see instructions below)
echo   2. Set your client_id in config.json
echo.
echo === Azure AD App Registration ===
echo  1. Go to https://portal.azure.com ^> Azure Active Directory ^> App registrations
echo  2. New registration: name = "Message Center Monitor", account type = Single tenant
echo  3. After creation, copy the Application (client) ID into config.json
echo  4. Go to API permissions ^> Add permission ^> Microsoft Graph ^> Delegated
echo  5. Add: ServiceMessage.Read.All
echo  6. Click "Grant admin consent for Cegeka"
echo  7. Under Authentication: Add platform = Mobile/desktop, redirect URI = https://login.microsoftonline.com/common/oauth2/nativeclient
echo.
pause
