@echo off
setlocal enabledelayedexpansion

REM ===============================
REM CONFIG
REM ===============================
set LOG_FILE=sync.log

echo ====================================== > %LOG_FILE%
echo Git Sync Started at %DATE% %TIME% >> %LOG_FILE%
echo ====================================== >> %LOG_FILE%

echo.
echo Synchronizing...
echo Progress: 5%%
echo Initializing... >> %LOG_FILE%

REM ===============================
REM 1. Detect current branch
REM ===============================
for /f %%i in ('git branch --show-current') do set CURRENT_BRANCH=%%i

if "%CURRENT_BRANCH%"=="" (
    echo ERROR: Cannot detect current git branch.
    echo ERROR: Cannot detect current git branch. >> %LOG_FILE%
    exit /b 1
)

echo Current branch: %CURRENT_BRANCH% >> %LOG_FILE%

REM ===============================
REM 2. Generate timestamp
REM ===============================
for /f %%i in ('powershell -Command "Get-Date -Format yyyyMMdd-HHmmss"') do set TS=%%i

REM ===============================
REM 3. Prepare backup branch
REM ===============================
set BACKUP_BRANCH=backup/origin-%CURRENT_BRANCH%-%TS%
echo Backup branch (cloud): %BACKUP_BRANCH% >> %LOG_FILE%

REM ===============================
REM 4. Fetch origin
REM ===============================
echo Progress: 20%% - Fetching origin...
echo Fetching origin... >> %LOG_FILE%
git fetch origin >> %LOG_FILE% 2>&1

if errorlevel 1 (
    echo ERROR: git fetch failed.
    echo ERROR: git fetch failed. >> %LOG_FILE%
    exit /b 1
)

REM ===============================
REM 5. Check local file changes
REM ===============================
echo Progress: 40%% - Checking local changes...
git diff --quiet
if errorlevel 1 (
    set HAS_LOCAL_FILE_CHANGES=1
    echo Local working tree: DIRTY >> %LOG_FILE%
) else (
    set HAS_LOCAL_FILE_CHANGES=0
    echo Local working tree: CLEAN >> %LOG_FILE%
)

REM ===============================
REM 6A. Local DIRTY -> prioritize LOCAL
REM ===============================
if %HAS_LOCAL_FILE_CHANGES% EQU 1 (
    echo Progress: 70%% - Syncing branches...
    echo Prioritize LOCAL changes >> %LOG_FILE%

    echo Creating cloud backup branch... >> %LOG_FILE%
    git branch %BACKUP_BRANCH% origin/%CURRENT_BRANCH% >> %LOG_FILE% 2>&1
    git push origin %BACKUP_BRANCH% >> %LOG_FILE% 2>&1

    echo Cloud backup branch pushed: origin/%BACKUP_BRANCH% >> %LOG_FILE%

    git add . >> %LOG_FILE% 2>&1
    git commit -m "Auto commit local changes (%TS%)" >> %LOG_FILE% 2>&1

    git push origin %CURRENT_BRANCH% --force-with-lease >> %LOG_FILE% 2>&1

    echo Force push completed >> %LOG_FILE%
)

REM ===============================
REM 6B. Local CLEAN -> prioritize ORIGIN
REM ===============================
if %HAS_LOCAL_FILE_CHANGES% EQU 0 (
    echo Progress: 70%% - Syncing branches...
    echo Prioritize ORIGIN changes >> %LOG_FILE%

    git pull origin %CURRENT_BRANCH% >> %LOG_FILE% 2>&1
    echo Pull completed >> %LOG_FILE%
)

REM ===============================
REM 7. Done
REM ===============================
echo Progress: 100%%
echo. 
echo Git synchronization completed successfully!
echo. 

echo ====================================== >> %LOG_FILE%
echo Sync completed at %DATE% %TIME% >> %LOG_FILE%
echo ====================================== >> %LOG_FILE%

choice /c YN /m "Press Y to view detailed logs, or press any other key to exit."

if errorlevel 2 exit /b 0
if errorlevel 1 (
    notepad %LOG_FILE%
)
