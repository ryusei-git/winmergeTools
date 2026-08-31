@echo off
rem ---------------------------------------------------------------------------
rem Launcher for WinMergeReportTool.
rem
rem  - The source is UTF-8, so it is compiled with "javac -encoding UTF-8".
rem    Running "java WinMergeReportTool.java" directly is only safe on JDK 18+,
rem    because older JDKs default the source encoding to MS932 on Japanese
rem    Windows and mangle the Japanese string literals.
rem  - The current directory stays as the caller's, which is what the tool
rem    compares.
rem  - Works with an installed JDK or a portable (zip) one; see tools\find_jdk.bat.
rem    The system PATH is never modified.
rem  - Console encoding is left alone on purpose: Java writes its output using
rem    the console's own code page, so Japanese shows correctly as-is.
rem
rem Usage:
rem   cd D:\tests
rem   C:\tools\winmergeTools\run.bat
rem   C:\tools\winmergeTools\run.bat --winmerge "C:\Program Files\WinMerge\WinMergeU.exe"
rem
rem NOTE: this file must stay ASCII-only (see tools\find_jdk.bat for why).
rem ---------------------------------------------------------------------------
setlocal

set "TOOLDIR=%~dp0"
set "CLASSDIR=%TEMP%\winmergeReportTool"

call "%TOOLDIR%tools\find_jdk.bat"
if errorlevel 1 exit /b 1

"%JAVAC%" -encoding UTF-8 -d "%CLASSDIR%" "%TOOLDIR%WinMergeReportTool.java"
if errorlevel 1 (
  echo [ERROR] Compilation failed.
  exit /b 1
)

"%JAVA%" -cp "%CLASSDIR%" WinMergeReportTool %*
exit /b %errorlevel%
