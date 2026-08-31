@echo off
rem ---------------------------------------------------------------------------
rem Prepare a Windows test environment and check that WinMerge behaves the way
rem WinMergeReportTool expects.
rem
rem   1. Resolve a JDK and compile
rem   2. Generate the sample data
rem   3. Locate WinMergeU.exe
rem   4. Smoke-test WinMerge (does /enableexitcode work, is /or written)
rem
rem Usage:
rem   tools\setup_test_env.bat              ... creates .\sample
rem   tools\setup_test_env.bat D:\tests     ... creates D:\tests\sample
rem
rem NOTE: this file must stay ASCII-only. cmd.exe reads batch files by byte
rem offset, so multibyte characters (and any chcp switch) can split lines in
rem the wrong place and corrupt parsing. Japanese output comes from the Java
rem programs, which encode it for the console correctly.
rem ---------------------------------------------------------------------------
setlocal

set "TOOLDIR=%~dp0"
set "SRCDIR=%TOOLDIR%.."
set "CLASSDIR=%TEMP%\wmrt"

set "BASEDIR=%~1"
if "%BASEDIR%"=="" set "BASEDIR=%CD%"
set "DEST=%BASEDIR%\sample"

echo ==================================================
echo  WinMergeReportTool - test environment setup
echo ==================================================
echo.

rem --- 1. JDK and compilation ----------------------------------------------
echo [1/4] Resolving a JDK and compiling
call "%TOOLDIR%find_jdk.bat"
if errorlevel 1 exit /b 1
"%JAVAC%" -encoding UTF-8 -d "%CLASSDIR%" "%SRCDIR%\WinMergeReportTool.java" "%TOOLDIR%MakeSampleData.java"
if errorlevel 1 (
  echo   [ERROR] Compilation failed.
  exit /b 1
)
echo   JDK     : %JAVAC%
echo   Classes : %CLASSDIR%
echo.

rem --- 2. Sample data ------------------------------------------------------
echo [2/4] Generating sample data in %DEST%
"%JAVA%" -cp "%CLASSDIR%" MakeSampleData --dest "%DEST%" --clean
if errorlevel 1 (
  echo   [ERROR] Failed to generate the sample data.
  exit /b 1
)
echo.

rem --- 3. Locate WinMergeU.exe ---------------------------------------------
echo [3/4] Looking for WinMergeU.exe
set "PF86=%ProgramFiles(x86)%"
set "WM="
if defined WINMERGE_PATH if exist "%WINMERGE_PATH%" set "WM=%WINMERGE_PATH%"
if not defined WM if exist "%ProgramFiles%\WinMerge\WinMergeU.exe" set "WM=%ProgramFiles%\WinMerge\WinMergeU.exe"
if not defined WM if exist "%PF86%\WinMerge\WinMergeU.exe" set "WM=%PF86%\WinMerge\WinMergeU.exe"
if not defined WM if exist "%LOCALAPPDATA%\Programs\WinMerge\WinMergeU.exe" set "WM=%LOCALAPPDATA%\Programs\WinMerge\WinMergeU.exe"
if not defined WM (
  echo   [ERROR] WinMergeU.exe not found.
  echo           Install WinMerge, then either set WINMERGE_PATH to the full
  echo           path of the exe, or pass --winmerge when running the tool.
  exit /b 1
)
echo   %WM%
echo.

rem --- 4. WinMerge smoke test ----------------------------------------------
echo [4/4] Smoke-testing WinMerge
set "SM=%TEMP%\wm_smoke"
if exist "%SM%" rd /s /q "%SM%"
mkdir "%SM%\left" "%SM%\right"
echo line1>"%SM%\left\a.txt"
echo line2>"%SM%\left\b.txt"
echo line1>"%SM%\right\a.txt"
echo LINE2-CHANGED>"%SM%\right\b.txt"
xcopy "%SM%\left" "%SM%\same\" /e /i /q >nul

echo   (a) folders that differ  ... expecting exit code 1 and an HTML report
"%WM%" /r /e /u /minimize /noninteractive /enableexitcode /cfg ReportFiles/ReportType=2 /or "%SM%\diff.html" "%SM%\left" "%SM%\right"
set "RC_DIFF=%ERRORLEVEL%"
echo       exit code = %RC_DIFF%
if exist "%SM%\diff.html" (echo       report    = written) else (echo       report    = NOT written)

echo   (b) folders that match   ... expecting exit code 0
"%WM%" /r /e /u /minimize /noninteractive /enableexitcode /cfg ReportFiles/ReportType=2 /or "%SM%\same.html" "%SM%\left" "%SM%\same"
set "RC_SAME=%ERRORLEVEL%"
echo       exit code = %RC_SAME%
echo.

echo ==================== RESULT ====================
if "%RC_DIFF%"=="1" if "%RC_SAME%"=="0" (
  echo   OK : /enableexitcode works. Run the tool as-is.
  goto :next
)
if "%RC_DIFF%"=="0" (
  echo   NG : the folders differ but WinMerge returned 0.
  echo        This WinMerge probably predates /enableexitcode ^(2.14 or older^).
  echo        Run WinMergeReportTool with --no-exitcode.
  goto :next
)
echo   ?? : unexpected exit codes ^(diff=%RC_DIFF%, same=%RC_SAME%^).
echo        Check the WinMerge version and the reports left in %SM%.

:next
echo ===============================================
echo.
echo Next:
echo   cd /d "%DEST%"
echo   "%SRCDIR%\run.bat" --winmerge "%WM%"
echo.
echo Reports are written to %DEST%\report
exit /b 0
