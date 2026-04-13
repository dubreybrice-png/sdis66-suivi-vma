@echo off
REM Helper script for Suivi VMA deployment
REM Usage: clasp-helper.bat [push|deploy|pull|open|push-deploy|sync|help]
REM
REM IMPORTANT: Always use "sync" to keep GitHub AND GAS in sync!

setlocal enabledelayedexpansion

set action=%1
if "%action%"=="" set action=help

if /i "%action%"=="sync" (
    echo ══════════════════════════════════════════════
    echo  SYNC: GitHub + Google Apps Script
    echo ══════════════════════════════════════════════
    echo.
    echo [1/4] Adding all changes to git...
    git add -A
    echo.
    echo [2/4] Committing to git...
    set /p MSG="Message de commit (ou Entree pour auto): "
    if "!MSG!"=="" set MSG=sync: update from %COMPUTERNAME% on %DATE% %TIME%
    git commit -m "!MSG!"
    if errorlevel 1 (
        echo [!] Nothing to commit or error — continuing anyway...
    )
    echo.
    echo [3/4] Pushing to GitHub...
    git push
    if errorlevel 1 (
        echo [!] Git push failed! Fix manually.
        pause
        exit /b 1
    )
    echo.
    echo [4/4] Pushing to Google Apps Script...
    cmd /c "clasp push --force"
    echo.
    echo ══════════════════════════════════════════════
    echo  [+] SYNC COMPLETE — GitHub + GAS are in sync!
    echo ══════════════════════════════════════════════
    pause
) else if /i "%action%"=="push" (
    echo [*] Pushing code to Google Apps Script...
    cmd /c "clasp push --force"
    echo [+] Push completed!
    echo.
    echo [!] WARNING: GAS updated but GitHub may be out of sync!
    echo [!] Consider using "clasp-helper.bat sync" instead.
    pause
) else if /i "%action%"=="deploy" (
    echo [*] Creating deployment...
    cmd /c "clasp deploy"
    echo [+] Deployment completed!
    pause
) else if /i "%action%"=="pull" (
    echo [*] Pulling code from Google Apps Script...
    cmd /c "clasp pull"
    echo [+] Pull completed!
    echo.
    echo [!] Don't forget to commit + push to GitHub after reviewing!
    echo [!] Run: clasp-helper.bat sync
    pause
) else if /i "%action%"=="push-deploy" (
    echo [*] Full sync + deploy...
    echo.
    REM First sync to GitHub
    git add -A
    set /p MSG="Message de commit (ou Entree pour auto): "
    if "!MSG!"=="" set MSG=sync+deploy from %COMPUTERNAME% on %DATE% %TIME%
    git commit -m "!MSG!"
    git push
    echo.
    echo [*] Pushing to GAS...
    cmd /c "clasp push --force"
    echo [+] Push completed!
    echo.
    echo [*] Creating new deployment...
    for /f "tokens=2 delims= " %%A in ('cmd /c "clasp deploy" 2^>^&1 ^| findstr /i "Deployed"') do set DEPLOY_ID=%%A
    echo [+] Deployment ID: !DEPLOY_ID!
    echo.
    echo [*] Storing public deploy URL...
    set DEPLOY_URL=https://script.google.com/macros/s/!DEPLOY_ID!/exec
    echo     URL: !DEPLOY_URL!
    cmd /c "clasp run setDeployUrl -p [\"!DEPLOY_URL!\"]" 2>nul
    if errorlevel 1 (
        echo [!] Could not auto-store URL via clasp run.
        echo [!] Run this manually in GAS editor console:
        echo     setDeployUrl("!DEPLOY_URL!")
    ) else (
        echo [+] URL stored in script properties!
    )
    echo.
    echo [+] GitHub + GAS synced AND deployed!
    pause
) else if /i "%action%"=="open" (
    echo [*] Opening script editor...
    cmd /c "clasp open"
    pause
) else (
    echo SDIS66 Suivi VMA - Clasp Helper
    echo ================================
    echo.
    echo Usage: clasp-helper.bat [command]
    echo.
    echo Commands:
    echo   sync          - [RECOMMENDED] Git commit+push + clasp push (keeps everything in sync)
    echo   push-deploy   - Sync + create new GAS deployment
    echo   pull          - Pull code from Google Apps Script
    echo   push          - Push to GAS only (WARNING: may desync GitHub)
    echo   deploy        - Create new GAS deployment only
    echo   open          - Open script in browser
    echo   help          - Show this help message
    echo.
    echo ALWAYS use "sync" or "push-deploy" to avoid losing features!
    echo.
    pause
)
