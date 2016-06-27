c:\Python34\Scripts\pyinstaller.exe -w -n scheduler -i sched.ico win32gui_taskbar.py

md dist\scheduler\conf

xcopy conf dist\scheduler\conf /y

xcopy status.dat dist\scheduler /y

xcopy readme.txt dist\scheduler /y

xcopy sched.ico dist\scheduler /y

