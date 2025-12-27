@echo off
setlocal EnableDelayedExpansion

REM ===============================
REM CONFIG
REM ===============================
set LOG_FILE=.log.log

REM ===============================
REM INIT LOG
REM ===============================
(
echo ======================================
echo Git Sync Started at %DATE% %TIME%
echo ======================================
) > "%LOG_FILE%"

echo.
echo Synchronizing...
echo Progress: 5%%
call :log INFO Initializing

REM ===============================
REM 1. Detect current branch
REM ===============================
for /f %%i in ('git branch --show-current') do set CURRENT_BRANCH=%%i

if "%CURRENT_BRANCH%"=="" (
    echo ERROR: Cannot detect current git branch.
    call :log ERROR Cannot detect current git branch
    exit /b 1
)

call :log INFO Current branch: %CURRENT_BRANCH%

REM ===============================
REM 2. Generate timestamp
REM ===============================
for /f %%i in ('powershell -Command "Get-Date -Format yyyyMMdd-HHmmss"') do set TS=%%i

REM ===============================
REM 3. Prepare backup branch
REM ===============================
set BACKUP_BRANCH=backup/origin-%CURRENT_BRANCH%-%TS%
call :log INFO Cloud backup branch: %BACKUP_BRANCH%

REM ===============================
REM 4. Fetch origin
REM ===============================
echo Progress: 20%% - Fetching origin...
call :log INFO Fetching origin
git fetch origin >> "%LOG_FILE%" 2>&1

if errorlevel 1 (
    echo ERROR: git fetch failed.
    call :log ERROR git fetch failed
    exit /b 1
)

REM ===============================
REM 5. Check local file changes
REM ===============================
echo Progress: 40%% - Checking local changes...
git diff --quiet

if errorlevel 1 (
    set HAS_LOCAL_FILE_CHANGES=1
    call :log INFO Local working tree: DIRTY
) else (
    set HAS_LOCAL_FILE_CHANGES=0
    call :log INFO Local working tree: CLEAN
)

REM ===============================
REM 6A. Local DIRTY -> prioritize LOCAL
REM ===============================
if %HAS_LOCAL_FILE_CHANGES% EQU 1 (
    echo Progress: 70%% - Syncing branches...
    call :log INFO Strategy: PRIORITIZE LOCAL

    call :log INFO Creating cloud backup branch
    git branch %BACKUP_BRANCH% origin/%CURRENT_BRANCH% >> "%LOG_FILE%" 2>&1
    git push origin %BACKUP_BRANCH% >> "%LOG_FILE%" 2>&1
    call :log INFO Backup pushed: origin/%BACKUP_BRANCH%

    git add . >> "%LOG_FILE%" 2>&1
    git commit -m "Auto commit local changes (%TS%)" >> "%LOG_FILE%" 2>&1

    git push origin %CURRENT_BRANCH% --force-with-lease >> "%LOG_FILE%" 2>&1
    call :log INFO Force push completed
)

REM ===============================
REM 6B. Local CLEAN -> prioritize ORIGIN
REM ===============================
if %HAS_LOCAL_FILE_CHANGES% EQU 0 (
    echo Progress: 70%% - Syncing branches...
    call :log INFO Strategy: PRIORITIZE ORIGIN

    git pull origin %CURRENT_BRANCH% >> "%LOG_FILE%" 2>&1
    call :log INFO Pull completed
)

REM ===============================
REM 7. Done
REM ===============================
echo Progress: 100%%
echo.
echo Git synchronization completed successfully!
echo.

(
echo ======================================
echo Sync completed at %DATE% %TIME%
echo ======================================
) >> "%LOG_FILE%"

choice /c YN /m "Press Y to view detailed logs, or press any other key to exit."

if errorlevel 2 exit /b 0
if errorlevel 1 notepad "%LOG_FILE%"

exit /b 0

REM ===============================
REM LOG FUNCTION
REM ===============================
:log
REM Usage: call :log LEVEL MESSAGE
echo [%1] %2 >> "%LOG_FILE%"
exit /b 0
