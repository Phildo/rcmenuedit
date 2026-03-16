@echo off
setlocal EnableExtensions EnableDelayedExpansion

set "ROOT=%~dp0"
set "SRC=%ROOT%src"
set "PACKAGE=%ROOT%package"
set "OUTPUT=%PACKAGE%\rcmenuedit.exe"

if not exist "%SRC%" (
    echo Source folder not found: "%SRC%"
    exit /b 1
)

if not exist "%SRC%\*.cpp" (
    echo No .cpp files found under "%SRC%"
    exit /b 1
)

if not exist "%PACKAGE%" mkdir "%PACKAGE%"

set "SOURCES="
for %%F in ("%SRC%\*.cpp") do (
    set "SOURCES=!SOURCES! "%%~fF""
)

clang++ -std=c++17 -O2 -Wall -Wextra -Wpedantic !SOURCES! -ladvapi32 -lshell32 -o "%OUTPUT%"
if errorlevel 1 exit /b %errorlevel%

echo Built "%OUTPUT%"
