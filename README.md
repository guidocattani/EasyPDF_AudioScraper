# EasyPDF_AudioScraper
This repository contains a program to analyse a PDF-file showing tone audiogram and speech audiogram graphs and to make a description of the hearing loss. 

Text from the PDF is analysed to describe the tone audiometry results and the test condition. Patient information are also retrieved from the text. Test results are analysed to classify the type and extent of the hearing loss.
The speech audiogram is analysed reading the colored (red and blue) pixels of the psychometric curve. The maximum score, the associated speech stimulus level and the Speech Recognition Thresholds (SRT) are calculated. Because of the graphic method, the results are sometimes not perfect, but reasonable adequate. 
The code is optimised for the EasyViewer layout used by Adelante Audiology & Communication (Hoensbroek, The Netherlands). 
In addition to the "readme" and "licence", the repository contains 4 files. 
The code was first written and debugged in Jupyter Lab, generating the .ipynb file. 
A .py script was downloaded from Jupyter Lab and tested, first in a Linux Mint 21 (64 bits) system and secondly in a Windows 10 (64 bits) system. 
An executable file .exe was generated in the Windows system using Pyinstaller. You can see the terminal commands used in the text file, pyinstaller_command.txt. The Python code uses some libraries: Colorama, Numpy, Pillow, Pymupdf, Os-sys and Regex. You have to install these libraries in your computer (along Python) to run the Python code or to generate an executable.
In Linux a binary file was successfully generated, but is unfortunately too large to be stored in this repository. 
The application EasyPDF_AudioScraper_v1.exe opens a screen and asks for the location of the file to analyse. The results are showed on screen and saved in the user’s download location. 
If you have questions or suggestions, feel free to contact me. Thank you for your attention and consideration. 

Guido Cattani

guidocattani@gmail.com
