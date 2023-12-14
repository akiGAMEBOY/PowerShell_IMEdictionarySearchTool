@ECHO OFF
@REM #################################################################################
@REM # 処理名　｜IMEdictionarySearchTool（起動用バッチ）
@REM # 機能　　｜PowerShell起動用のバッチ
@REM #--------------------------------------------------------------------------------
@REM # 　　　　｜-
@REM #################################################################################
ECHO *---------------------------------------------------------
ECHO *
ECHO *  IMEdictionarySearchTool
ECHO *
ECHO *---------------------------------------------------------
ECHO.
ECHO.
SET RETURNCODE=0
@REM PowerShell Core インストール確認
WHERE /Q pwsh
IF %ERRORLEVEL% == 0 (
    @REM PowerShell Core で実行する場合
    pwsh -NoProfile -ExecutionPolicy Unrestricted -File .\source\powershell\Main.ps1
) ELSE (
    @REM PowerShell 5.1  で実行する場合
    powershell -NoProfile -ExecutionPolicy Unrestricted -File .\source\powershell\Main.ps1
)

SET RETURNCODE=%ERRORLEVEL%

ECHO.
ECHO 処理が終了しました。
ECHO いずれかのキーを押すとウィンドウが閉じます。
PAUSE > NUL
EXIT %RETURNCODE%
