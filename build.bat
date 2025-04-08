@echo off
echo Installing dependencies...
pip install -r requirements.txt

echo Creating executable...
pyinstaller --noconfirm --onefile --windowed ^
    --add-data "requirements.txt;." ^
    --icon=NONE ^
    --name "FileCompressor7z" ^
    compressapppy/file_compressor_7z.py

echo Done! The executable is in the dist folder.
pause 