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
REM 5. Check LOCAL FILE CHANGES (WORKING TREE)
REM ======================================
git diff --quiet
if errorlevel 1 (
    set HAS_LOCAL_FILE_CHANGES=1
) else (
    set HAS_LOCAL_FILE_CHANGES=0
)

REM ======================================
REM 6A. CASE: Local has FILE CHANGES
REM ======================================
if %HAS_LOCAL_FILE_CHANGES% EQU 1 (
    echo Local working tree is DIRTY
    echo -> Prioritize LOCAL
    echo -> Backup origin before force

    git branch %BACKUP_BRANCH% origin/%CURRENT_BRANCH%
    git push origin %BACKUP_BRANCH%

    git add .
    git commit -m "Auto commit local file changes (%TS%)"

    git push origin %CURRENT_BRANCH% --force-with-lease

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

    echo PULL SYNC DONE
)

REM ======================================
REM 7. Done
REM ======================================
echo --------------------------------------
echo SYNC COMPLETED SUCCESSFULLY
echo --------------------------------------
pause
