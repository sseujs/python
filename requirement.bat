curl -o python-installer.exe https://www.python.org/ftp/python/3.12.1/python-3.12.1-amd64.exe
python-installer.exe
python -m pip install --upgrade pip
pip install -r text.txt
del /f /q text.txt
del /f /q readme.md
del /f /q license
ping localhost -n 5 > nul
del /f /q requirement.bat
continue
