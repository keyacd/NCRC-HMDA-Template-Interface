# NCRC HMDA Template Interface (version 4.0)
- A program that automates the process of filling out the NCRC HMDA Template.
- *Created By:* Darien Keyack, for [NCRC] (http://www.ncrc.org/).
- *Last Updated:* 8/18/2016

# Instructions
- Go to the [github repository] (https://github.com/keyacd/NCRC-HMDA-Template-Interface), select *Clone or Download*, and then *Download ZIP*.
  - Once the files are downloaded, unzip them.
- Go to the [CFPB website] (http://www.consumerfinance.gov/data-research/hmda/explore) and download the HMDA data you want to use. Do not filter by lender; the program will do that for you later. Apply any and all other filters you want to use. 
  - Once your data is downloaded, move it into the folder *Template_Interface_v3*, in the same folder as the files you downloaded and unzipped.
- In the folder *Template_Interface_v3*, find the Application file named *Template_Interface_v3* and launch it. 
  - Alternatively, you can just click *Template_Interface_v3 - Shortcut* in the parent folder of *Template_Interface_v3* .
- Follow all instructions provided within the program.
- Once the program is finished filling out the template for you, it will return to the home screen. 
  - Be sure to save the completed template, if you want to keep it.

# Programs used
- Python 3.4.3
- pywin32 220
- PyQt4 11.4
- xlwings 0.9.2
- *DemDataMSA.xlsx* and *HMDA Template u.v.12.xltm* were provided by NCRC.

# Distribution
- PyInstaller 3.2 was used to make the .exe and distribution folder. It was run on a computer with Windows 7 Professional (x64).
- Note for any future attempts: the folder will not automatically contain all the files needed for PyQt. You need to create a folder named *platforms* in qt4_plugins (in the dist), then go copy *qwindows.dll* from *C:\Python34\Lib\site-packages\PyQt4\plugins\platforms* to the *platforms* folder you created.
  - You will also need to make sure to include *--windowed* when running PyInstaller.
