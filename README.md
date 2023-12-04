# EasyPDF_AudioScraper
This repository contains a program to analyse a PDF-file containing toneaudiogram and speechaudiogram graphs and to make a description of the hearing loss. 
Text from the PDF is analysed to describe de test and the toneaudiometry results. Patient information are retrieved. Test results are analysed to classify the type and extent of the hearing loss.
The speechaudiogram is analysed reading the colored (red and blue) pixels of the psychometric curve. The maximum score, the correspondent speech stimulus level and the Speech Recognition Thresholds are calculated. 
Because of the graphic method the results are sometimes not perfect, but reasonable adequate. 
The repository contains 4 files. 
The code was written and debugged in Jupyter Lab, generating the .ipynb file. 
A .py script was downloaded from Jupyter Lab and tested, first in a Linux Mint 20 (64 bits) system and secondly in Windows 10 (64 bits). 
The python code use some libraries: Colorama,  Numpy, Pillow, Pymupdf, Os-sys and Regex.
An executable file .exe was generated in de Windows system using Pyinstaller. You can see the terminal commands used. 
In Linux a binary file was generated, This file is unfortunately to large to be stored in this repository. 
If you have questions or suggestions, feel free to contact me. Thank you for your attention. 

guidocattani@gmail.com
