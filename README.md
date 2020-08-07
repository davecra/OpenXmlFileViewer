# Office Open XML File Viewer
All Office Open XML files are really just a ZIP Package, renamed to .pptx, .xlsx, .docx, etc. This file viewer allows you to peek inside one of these files to read or modify the individual parts of the file.

**NOTE**: I created this application for myself back in 2013 and forgot about it. I recently needed it again for a project I was working on. I did a housecleaning, converted it to x64, and published it here.

## Install
To install this application, please download it from here:
https://github.com/davecra/OpenXmlFileViewer/raw/master/OpenXmlFileViewer_1.0.0.1.zip

Unzip this file and place it in any folder. It is suggested to place it in c:\Program Files\OpenXmlFileViewer\OpenXmlFileViewer.exe

## Usage
The first time you use this for each file type you wich to open, you will need to perform these steps:

1) Hold the SHIFT key and then right-click on the file you wish to view.
2) Click **Open With** and then select **Choose another app**
3) Click **More apps**, then scroll to the bottom of the list and select **Look for another app on this PC**
4) Browse to *OpenXmlFileViewer.exe* (if you installed it in Program Files, per directions above, select it from *c:\Program Files\OpenXmlFileViewer\*).
5) You will now see "OpenXmlFileViewer" in the list. DO NOT CLICK TO CHECK: **Always use this app to open XXXX files** or you will disassociate Office from the file.
6) The file will now open in the Office Open XML File Viewer.

You will have to do this one time for each file type:
* Word - .docx, .docm
* Excel - .xlsx, .xlsm
* PowerPoint - .pptx, .pptm

## Reporting issues
Please submit any issues you have with this to the Issues section above.
