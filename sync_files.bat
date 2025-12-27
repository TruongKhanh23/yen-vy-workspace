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
REM 2. Generate timestamp for backup
REM ======================================
for /f %%i in ('powershell -Command "Get-Date -Format yyyyMMdd-HHmmss"') do set TS=%%i

REM ======================================
REM 3. Prepare backup branch name
REM ======================================
set BACKUP_BRANCH=backup/origin-%CURRENT_BRANCH%-%TS%

echo --------------------------------------
echo Current branch : %CURRENT_BRANCH%
echo Backup branch  : %BACKUP_BRANCH%
echo --------------------------------------

REM ======================================
REM 4. Fetch latest origin state
REM ======================================
git fetch origin
if errorlevel 1 (
    echo ERROR: git fetch failed.
    exit /b 1
)

REM ======================================
REM 5. Backup origin branch BEFORE force
REM ======================================
echo Creating backup branch from origin/%CURRENT_BRANCH% ...
git branch %BACKUP_BRANCH% origin/%CURRENT_BRANCH%
if errorlevel 1 (
    echo ERROR: Failed to create backup branch.
    exit /b 1
)

git push origin %BACKUP_BRANCH%
if errorlevel 1 (
    echo ERROR: Failed to push backup branch to origin.
    exit /b 1
)

REM ======================================
REM 6. Check LOCAL FILE CHANGES (working tree)
REM ======================================
git status --porcelain > __git_status.tmp

set HAS_LOCAL_FILE_CHANGES=0
for %%A in (__git_status.tmp) do set HAS_LOCAL_FILE_CHANGES=1
del __git_status.tmp

REM ======================================
REM 7. Commit ONLY if local file changes exist
REM ======================================
if %HAS_LOCAL_FILE_CHANGES% EQU 1 (
    echo Local working tree is DIRTY -> committing file changes
    git add .
    git commit -m "Auto commit local file changes (%TS%)"
) else (
    echo Local working tree is CLEAN -> skip commit
)

REM ======================================
REM 8. Force push local to origin (SAFE)
REM ======================================
echo Force pushing local branch to origin...
git push origin %CURRENT_BRANCH% --force-with-lease
if errorlevel 1 (
    echo ERROR: Force push failed.
    exit /b 1
)

REM ======================================
REM 9. Done
REM ======================================
echo --------------------------------------
echo FORCE SYNC COMPLETED SUCCESSFULLY
echo Backup branch: origin/%BACKUP_BRANCH%
echo --------------------------------------

pause
