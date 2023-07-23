python ./setupfile.py
SET cwd=%CD% || echo failed to set cwd
cd %userprofile%\Desktop || echo failed to change directory
mklink GSTscraper %cwd%\gst_gui.exe || echo failed to create link