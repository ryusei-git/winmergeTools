@echo off
rem ---------------------------------------------------------------------------
rem Resolve a JDK and expose it as JAVAC / JAVA to the calling batch file.
rem Called from run.bat and setup_test_env.bat.
rem
rem Search order:
rem   1. javac on PATH (normal installation)
rem   2. %JAVA_HOME%\bin\javac.exe
rem   3. a "jdk" folder next to the tool (portable JDK extracted beside it)
rem   4. C:\tools\jdk, C:\jdk, D:\tools\jdk, and jdk* under those roots
rem
rem Nothing is written to the system PATH or the registry. The variables live
rem only inside the caller's setlocal scope.
rem
rem NOTE: this file must stay ASCII-only. cmd.exe reads batch files by byte
rem offset, so multibyte characters (and any chcp switch) can split lines in
rem the wrong place and corrupt parsing.
rem ---------------------------------------------------------------------------

set "JAVAC="
set "JAVA="

where javac >nul 2>&1
if not errorlevel 1 (
  set "JAVAC=javac"
  set "JAVA=java"
  goto :found
)

if defined JAVA_HOME if exist "%JAVA_HOME%\bin\javac.exe" (
  set "JAVAC=%JAVA_HOME%\bin\javac.exe"
  set "JAVA=%JAVA_HOME%\bin\java.exe"
  goto :found
)

for %%D in ("%~dp0..\jdk" "%~dp0jdk" "C:\tools\jdk" "C:\jdk" "D:\tools\jdk") do (
  if exist "%%~fD\bin\javac.exe" (
    set "JAVAC=%%~fD\bin\javac.exe"
    set "JAVA=%%~fD\bin\java.exe"
    goto :found
  )
)

rem Also accept the extraction folder name as-is, e.g. jdk-21.0.5+11
for /d %%D in ("%~dp0..\jdk*" "C:\tools\jdk*" "C:\jdk*" "D:\tools\jdk*") do (
  if exist "%%~fD\bin\javac.exe" (
    set "JAVAC=%%~fD\bin\javac.exe"
    set "JAVA=%%~fD\bin\java.exe"
    goto :found
  )
)

echo [ERROR] No JDK found.
echo.
echo   If you do not want to install one, extract a portable JDK (.zip) to
echo   one of these locations - no installer, no registry, delete to remove:
echo.
echo       C:\tools\jdk\bin\javac.exe
echo       %~dp0..\jdk\bin\javac.exe
echo.
echo   Or point JAVA_HOME at it in this command prompt:
echo       set "JAVA_HOME=C:\tools\jdk"
echo.
echo   Download: https://adoptium.net/temurin/releases/
echo       OS=Windows, Architecture=x64, Package Type=JDK, file type .zip
exit /b 1

:found
exit /b 0
