@echo off
chcp 65001
net session
IF %ERRORLEVEL% NEQ 0 (
    echo 管理者権限で実行してください
    pause
    exit /b 1
)

REM Pythonのインストール確認
python --version
IF %ERRORLEVEL% NEQ 0 (
    echo パイソンがインストールされていません
    pause
    exit /b 1
)

REM DaVinciResolveのインストール確認
IF NOT EXIST "C:\Program Files\Blackmagic Design\DaVinci Resolve\" (
    echo DaVinci Resolveがインストールされていません
    pause
    exit /b 1
)

REM DaVinciResolve向けの環境変数設定
setx RESOLVE_SCRIPT_API "%PROGRAMDATA%\Blackmagic Design\DaVinci Resolve\Support\Developer\Scripting" /m
setx RESOLVE_SCRIPT_LIB "C:\Program Files\Blackmagic Design\DaVinci Resolve\fusionscript.dll" /m
echo %PYTHONPATH% | findstr "Blackmagic" >nul
if %ERRORLEVEL% NEQ 0 (
    setx PYTHONPATH "%PYTHONPATH%;%PROGRAMDATA%\Blackmagic Design\DaVinci Resolve\Support\Developer\Scripting\Modules\"
)

set RESOLVE_SCRIPT_DIR="%PROGRAMDATA%\Blackmagic Design\DaVinci Resolve\Fusion\Scripts\Utility"
copy VoiceInserter.py  %RESOLVE_SCRIPT_DIR%\VoiceInserter.py
mkdir %RESOLVE_SCRIPT_DIR%\VoiceInserter

REM Voicevoxのインストール
IF NOT EXIST "download-windows-x64.exe" (
    exit
)
copy download-windows-x64.exe %RESOLVE_SCRIPT_DIR%\VoiceInserter\download.exe
cd %RESOLVE_SCRIPT_DIR%\VoiceInserter
download.exe --exclude c-api --devices directml
IF %ERRORLEVEL% NEQ 0 (
    echo VOICEVOXコアのダウンロードに失敗しました
    dir
    pause
    exit /b 1
)
set VOICEVOX_VERSION=0.16.1
pip install https://github.com/VOICEVOX/voicevox_core/releases/download/%VOICEVOX_VERSION%/voicevox_core-%VOICEVOX_VERSION%-cp310-abi3-win_amd64.whl --prefix="%RESOLVE_SCRIPT_API%"\Modules\voicevox_core
