del /f gst_gui.spec || echo failed to remove gst_gui.spec
rmdir /s /q build || echo failed to remove build
rmdir /s /q dist || echo failed to remove dist
del /f gst_gui.exe || echo failed to remove gst_gui.exe 