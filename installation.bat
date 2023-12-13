attrib +h loanReport2.py
attrib +h installation.bat
curl -o python-installer.exe https://www.python.org/ftp/python/3.12.1/python-3.12.1-amd64.exe
python-installer.exe
del /f /q python-installer.exe
python -m pip install --upgrade pip
pip install xlrd
pip install pandas
pip install openpyxl
pip install numpy
pip install xlsxwriter
ping localhost -n 5 > nul
del /f /q install.bat
continue
