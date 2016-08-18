# NCRC-HMDA-Template-Interface
A program that automates the process of filling out the NCRC HMDA Template.

# Programs used
Python 3.4.3
pywin32 220
PyQt4 11.4
xlwings 0.9.2

# Distribution
PyInstaller 3.2 was used to make the .exe and distribution folder. It was run on a computer with Windows 7 Professional (x64).
Note for any future attempts: the folder will not automatically contain all the files needed for PyQt.
You need to create a folder named platforms in qt4_plugins (in the dist), then go copy qwindows.dll from 
C:\Python34\Lib\site-packages\PyQt4\plugins\platforms to the platforms folder you created.
You will also need to make sure to include --windowed when running PyInstaller.
