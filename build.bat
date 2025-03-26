@echo off
echo 音声文字起こしツールのビルドを開始します...

REM 必要なパッケージをインストール
echo 必要なパッケージをインストールしています...
pip install -r requirements.txt

REM FFmpegをダウンロード（存在しない場合）
if not exist ffmpeg (
    echo FFmpegをダウンロードしています...
    powershell -Command "& {Invoke-WebRequest -Uri 'https://github.com/BtbN/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win64-gpl.zip' -OutFile 'ffmpeg.zip'}"
    powershell -Command "& {Expand-Archive -Path 'ffmpeg.zip' -DestinationPath 'temp'}"
    mkdir ffmpeg
    mkdir ffmpeg\bin
    copy temp\ffmpeg-master-latest-win64-gpl\bin\*.* ffmpeg\bin\
    rmdir /s /q temp
    del ffmpeg.zip
)

REM アイコンファイルを作成（存在しない場合）
if not exist app.ico (
    echo アイコンファイルを作成しています...
    powershell -Command "& {Invoke-WebRequest -Uri 'https://www.google.com/favicon.ico' -OutFile 'app.ico'}"
)

REM PyInstallerでビルド
echo exeファイルをビルドしています...
pyinstaller --clean transcribe.spec

echo ビルドが完了しました。
echo 実行ファイルは dist フォルダ内に作成されています。
pause 