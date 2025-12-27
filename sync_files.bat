@echo off
setlocal enabledelayedexpansion

for /f %%i in ('git branch --show-current') do set BRANCH=%%i
for /f %%i in ('powershell -Command "Get-Date -Format yyyyMMdd-HHmmss"') do set TS=%%i

set BACKUP_BRANCH=backup/origin-%BRANCH%-%TS%

echo Current branch: %BRANCH%
echo Backup branch: %BACKUP_BRANCH%

git fetch origin

echo Creating backup branch...
git branch %BACKUP_BRANCH% origin/%BRANCH%
git push origin %BACKUP_BRANCH%

git status --porcelain > status.tmp
set SIZE=0
for %%A in (status.tmp) do set SIZE=%%~zA
del status.tmp

if %SIZE% NEQ 0 (
    git add .
    git commit -m "Auto commit before force sync (%TS%)"
)

echo Force pushing local to origin...
git push origin %BRANCH% --force-with-lease

echo DONE
pause
