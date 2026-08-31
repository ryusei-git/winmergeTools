@echo off
chcp 65001 >nul
rem ---------------------------------------------------------------------------
rem Windows のテスト環境を用意し、WinMerge が期待どおり動くかを確認する。
rem
rem   1. JDK の確認とコンパイル
rem   2. サンプルデータの生成
rem   3. WinMergeU.exe の場所を特定
rem   4. WinMerge の疎通確認（/enableexitcode と /or が効くか）
rem
rem 使い方:
rem   tools\setup_test_env.bat              ... カレントに sample フォルダを作る
rem   tools\setup_test_env.bat D:\tests     ... D:\tests\sample を作る
rem ---------------------------------------------------------------------------
setlocal

set "TOOLDIR=%~dp0"
set "SRCDIR=%TOOLDIR%.."
set "CLASSDIR=%TEMP%\wmrt"

set "BASEDIR=%~1"
if "%BASEDIR%"=="" set "BASEDIR=%CD%"
set "DEST=%BASEDIR%\sample"

echo ==================================================
echo  WinMergeReportTool テスト環境セットアップ
echo ==================================================
echo.

rem --- 1. JDK の確認とコンパイル -------------------------------------------
echo [1/4] JDK を確認してコンパイルします
where javac >nul 2>&1
if errorlevel 1 (
  echo   [ERROR] javac が見つかりません。JDK 11 以降をインストールし、PATH を通してください。
  exit /b 1
)
javac -version
javac -encoding UTF-8 -d "%CLASSDIR%" "%SRCDIR%\WinMergeReportTool.java" "%TOOLDIR%MakeSampleData.java"
if errorlevel 1 (
  echo   [ERROR] コンパイルに失敗しました。
  exit /b 1
)
echo   コンパイル先: %CLASSDIR%
echo.

rem --- 2. サンプルデータの生成 ---------------------------------------------
echo [2/4] サンプルデータを生成します: %DEST%
java -Dfile.encoding=UTF-8 -Dstdout.encoding=UTF-8 -cp "%CLASSDIR%" MakeSampleData --dest "%DEST%" --clean
if errorlevel 1 (
  echo   [ERROR] サンプルデータの生成に失敗しました。
  exit /b 1
)
echo.

rem --- 3. WinMergeU.exe の場所 ---------------------------------------------
echo [3/4] WinMergeU.exe を探します
set "PF86=%ProgramFiles(x86)%"
set "WM="
if defined WINMERGE_PATH if exist "%WINMERGE_PATH%" set "WM=%WINMERGE_PATH%"
if not defined WM if exist "%ProgramFiles%\WinMerge\WinMergeU.exe" set "WM=%ProgramFiles%\WinMerge\WinMergeU.exe"
if not defined WM if exist "%PF86%\WinMerge\WinMergeU.exe" set "WM=%PF86%\WinMerge\WinMergeU.exe"
if not defined WM if exist "%LOCALAPPDATA%\Programs\WinMerge\WinMergeU.exe" set "WM=%LOCALAPPDATA%\Programs\WinMerge\WinMergeU.exe"
if not defined WM (
  echo   [ERROR] WinMergeU.exe が見つかりません。
  echo           インストール後、環境変数 WINMERGE_PATH に exe のフルパスを設定するか、
  echo           WinMergeReportTool 実行時に --winmerge で指定してください。
  exit /b 1
)
echo   %WM%
echo.

rem --- 4. WinMerge の疎通確認 ----------------------------------------------
echo [4/4] WinMerge の疎通確認をします
set "SM=%TEMP%\wm_smoke"
if exist "%SM%" rd /s /q "%SM%"
mkdir "%SM%\left" "%SM%\right"
echo line1>"%SM%\left\a.txt"
echo line2>"%SM%\left\b.txt"
echo line1>"%SM%\right\a.txt"
echo LINE2-CHANGED>"%SM%\right\b.txt"
xcopy "%SM%\left" "%SM%\same\" /e /i /q >nul

echo   (a) 差分ありのフォルダ比較 ... 終了コード 1 と HTML 生成を期待
"%WM%" /r /e /u /minimize /noninteractive /enableexitcode /cfg ReportFiles/ReportType=2 /or "%SM%\diff.html" "%SM%\left" "%SM%\right"
set "RC_DIFF=%ERRORLEVEL%"
echo       終了コード = %RC_DIFF%
if exist "%SM%\diff.html" (echo       レポート     = 生成されました) else (echo       レポート     = 生成されませんでした)

echo   (b) 一致するフォルダ比較 ... 終了コード 0 を期待
"%WM%" /r /e /u /minimize /noninteractive /enableexitcode /cfg ReportFiles/ReportType=2 /or "%SM%\same.html" "%SM%\left" "%SM%\same"
set "RC_SAME=%ERRORLEVEL%"
echo       終了コード = %RC_SAME%
echo.

echo ==================== 判定 ====================
if "%RC_DIFF%"=="1" if "%RC_SAME%"=="0" (
  echo   OK : /enableexitcode が有効です。そのまま実行できます。
  goto :next
)
if "%RC_DIFF%"=="0" (
  echo   NG : 差分があるのに終了コードが 0 でした。
  echo        WinMerge が /enableexitcode に対応していない可能性があります^(2.14 以前^)。
  echo        WinMergeReportTool を --no-exitcode 付きで実行してください。
  goto :next
)
echo   ?? : 想定外の終了コードです ^(diff=%RC_DIFF%, same=%RC_SAME%^)。
echo        WinMerge のバージョンと、%SM% に生成されたレポートを確認してください。

:next
echo ==============================================
echo.
echo 次の手順:
echo   cd /d "%DEST%"
echo   "%SRCDIR%\run.bat" --winmerge "%WM%"
echo.
echo 生成された比較レポートは %DEST%\report に出力されます。
exit /b 0
