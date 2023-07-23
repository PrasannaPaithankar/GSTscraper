import os
import shutil 
os.system("pip install -r requirements.txt")
os.system("pyinstaller --onefile --windowed gst_gui.py")
shutil.move("dist/gst_gui.exe", "GSTscraper.exe")