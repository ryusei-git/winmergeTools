@echo off
rem ---------------------------------------------------------------------------
rem JDK の場所を解決し、環境変数 JAVAC / JAVA に設定する。
rem run.bat と setup_test_env.bat から call されるヘルパー。
rem
rem 探索順:
rem   1. PATH 上の javac（通常インストール）
rem   2. %JAVA_HOME%\bin\javac.exe
rem   3. ツールと同じ場所の jdk フォルダ（ポータブル JDK を隣に置いた場合）
rem   4. C:\tools\jdk, C:\jdk, D:\tools\jdk（よく使う展開先）
rem
rem システムの PATH やレジストリは変更しない。呼び出し元の setlocal の中でのみ有効。
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

rem 展開直後のフォルダ名（jdk-21.0.5+11 など）のまま置かれている場合も拾う
for /d %%D in ("%~dp0..\jdk*" "C:\tools\jdk*") do (
  if exist "%%~fD\bin\javac.exe" (
    set "JAVAC=%%~fD\bin\javac.exe"
    set "JAVA=%%~fD\bin\java.exe"
    goto :found
  )
)

echo [ERROR] JDK が見つかりません。
echo         インストールしたくない場合は、ポータブル JDK の zip を展開して
echo         次のいずれかに置いてください（インストーラ不要・フォルダ削除で撤去できます）。
echo.
echo           C:\tools\jdk\bin\javac.exe
echo           %~dp0..\jdk\bin\javac.exe
echo.
echo         または、実行するコマンドプロンプトで JAVA_HOME を指定してください。
echo           set "JAVA_HOME=C:\tools\jdk"
echo.
echo         zip 版の入手先: https://adoptium.net/temurin/releases/
echo           OS=Windows / Architecture=x64 / Package Type=JDK / 拡張子 .zip を選択
exit /b 1

:found
exit /b 0
