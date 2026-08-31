@echo off
chcp 65001 >nul
rem ---------------------------------------------------------------------------
rem WinMergeReportTool 起動用バッチ
rem
rem  - ソースは UTF-8 のため javac -encoding UTF-8 でコンパイルする
rem    （JDK 17 以前は既定の文字コードが MS932 になり、直接 java Xxx.java すると
rem      日本語リテラルが壊れるため）
rem  - コンソールを UTF-8 にして日本語の出力が化けないようにする
rem  - カレントディレクトリは呼び出し元のまま（＝比較対象のルートになる）
rem  - JDK はインストール版でもポータブル版（zip 展開）でもよい。
rem    探索順は tools\find_jdk.bat を参照。システムの PATH は変更しない。
rem
rem 使い方:
rem   cd D:\tests
rem   C:\tools\winmergeTools\run.bat
rem   C:\tools\winmergeTools\run.bat --winmerge "C:\Program Files\WinMerge\WinMergeU.exe"
rem ---------------------------------------------------------------------------
setlocal

set "TOOLDIR=%~dp0"
set "CLASSDIR=%TEMP%\winmergeReportTool"

call "%TOOLDIR%tools\find_jdk.bat"
if errorlevel 1 exit /b 1

"%JAVAC%" -encoding UTF-8 -d "%CLASSDIR%" "%TOOLDIR%WinMergeReportTool.java"
if errorlevel 1 (
  echo [ERROR] コンパイルに失敗しました。
  exit /b 1
)

"%JAVA%" -Dfile.encoding=UTF-8 -Dstdout.encoding=UTF-8 -Dstderr.encoding=UTF-8 ^
     -cp "%CLASSDIR%" WinMergeReportTool %*
exit /b %errorlevel%
