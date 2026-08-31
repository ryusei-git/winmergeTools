@echo off
rem ---------------------------------------------------------------------------
rem WinMergeReportTool 起動用バッチ（任意・無くても java コマンドだけで実行可能）
rem
rem  - ソースは UTF-8 のため javac -encoding UTF-8 でコンパイルする
rem    （JDK 17 以前は既定の文字コードが MS932 になり、直接 java Xxx.java すると
rem      日本語リテラルが壊れるため）
rem  - コンソールを UTF-8 にして日本語の出力が化けないようにする
rem  - カレントディレクトリは呼び出し元のまま（＝比較対象のルートになる）
rem
rem 使い方:
rem   cd D:\tests
rem   C:\tools\winmergeTools\run.bat
rem   C:\tools\winmergeTools\run.bat --winmerge "C:\Program Files\WinMerge\WinMergeU.exe"
rem ---------------------------------------------------------------------------
setlocal
chcp 65001 >nul

set "TOOLDIR=%~dp0"
set "CLASSDIR=%TEMP%\winmergeReportTool"

javac -encoding UTF-8 -d "%CLASSDIR%" "%TOOLDIR%WinMergeReportTool.java"
if errorlevel 1 (
  echo [ERROR] コンパイルに失敗しました。JDK 11 以降がインストールされ、javac に PATH が通っているか確認してください。
  exit /b 1
)

java -Dfile.encoding=UTF-8 -Dstdout.encoding=UTF-8 -Dstderr.encoding=UTF-8 ^
     -cp "%CLASSDIR%" WinMergeReportTool %*
exit /b %errorlevel%
