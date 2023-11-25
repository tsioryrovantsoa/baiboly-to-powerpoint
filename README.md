# baiboly-to-powerpoint
Mamadika ny Baiboly ho lasa PowerPoint

# dependance
altgraph==0.17.4
lxml==4.9.3
packaging==23.2
pefile==2023.2.7
Pillow==10.1.0
pyinstaller==6.2.0
pyinstaller-hooks-contrib==2023.10
python-dotenv==1.0.0
python-pptx==0.6.23
pywin32-ctypes==0.2.2
XlsxWriter==3.1.9

# get dependance
pip freeze > requirements.txt

# deploiement
pyinstaller --onefile --windowed --icon=baiboly.ico --distpath . main.py

pip install -r requirements.txt
