@echo off
setlocal enabledelayedexpansion

REM ======================================
REM 1. Detect current branch
REM ======================================
for /f %%i in ('git branch --show-current') do set CURRENT_BRANCH=%%i

if "%CURRENT_BRANCH%"=="" (
    echo ERROR: Cannot detect current git branch.
    exit /b 1
)

REM ======================================
REM 2. Generate timestamp
REM ======================================
for /f %%i in ('powershell -Command "Get-Date -Format yyyyMMdd-HHmmss"') do set TS=%%i

REM ======================================
REM 3. Prepare backup branch name
REM ======================================
set BACKUP_BRANCH=backup/origin-%CURRENT_BRANCH%-%TS%

echo --------------------------------------
echo Current branch : %CURRENT_BRANCH%
echo --------------------------------------

REM ======================================
REM 4. Fetch origin
REM ======================================
git fetch origin
if errorlevel 1 (
    echo ERROR: git fetch failed.
    exit /b 1
)

REM ======================================
REM 5. Check LOCAL FILE CHANGES (working tree)
REM ======================================
git status --porcelain > __git_status.tmp

set HAS_LOCAL_FILE_CHANGES=0
for %%A in (__git_status.tmp) do set HAS_LOCAL_FILE_CHANGES=1
del __git_status.tmp

REM ======================================
REM 6A. CASE: Local has FILE CHANGES
REM ======================================
if %HAS_LOCAL_FILE_CHANGES% EQU 1 (
    echo Local working tree is DIRTY
    echo -> Prioritize LOCAL
    echo -> Backup origin before force

    REM Backup origin branch
    git branch %BACKUP_BRANCH% origin/%CURRENT_BRANCH%
    if errorlevel 1 (
        echo ERROR: Failed to create backup branch.
        exit /b 1
    )

    git push origin %BACKUP_BRANCH%
    if errorlevel 1 (
        echo ERROR: Failed to push backup branch.
        exit /b 1
    )

    REM Commit local file changes
    git add .
    git commit -m "Auto commit local file changes (%TS%)"

    REM Force push local to origin (safe)
    git push origin %CURRENT_BRANCH% --force-with-lease
    if errorlevel 1 (
        echo ERROR: Force push failed.
        exit /b 1
    )

    echo FORCE SYNC DONE
    echo Backup branch: origin/%BACKUP_BRANCH%
)

REM ======================================
REM 6B. CASE: Local is CLEAN
REM ======================================
if %HAS_LOCAL_FILE_CHANGES% EQU 0 (
    echo Local working tree is CLEAN
    echo -> Prioritize ORIGIN
    echo -> Pull to sync origin into local

    git pull origin %CURRENT_BRANCH%
    if errorlevel 1 (
        echo ERROR: git pull failed.
        exit /b 1
    )

    echo PULL SYNC DONE
)

REM ======================================
REM 7. Done
REM ======================================
echo --------------------------------------
echo SYNC COMPLETED SUCCESSFULLY
echo --------------------------------------
pause
