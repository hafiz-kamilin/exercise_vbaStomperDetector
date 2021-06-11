# VBA Stomper Detector

## Introduction

<p align = "center">
  <img src = "https://raw.githubusercontent.com/hafiz-kamilin/exercise_vbaStomperDetector/main/source/flowchart.png" width = "395" height = "653"/>
</p>

This is an example concept that shows how we can detect the disparity in compiled VBA macro (known as p-code) and VBA macro source code in Microsoft Office file. This disparity reveal that the Microsoft Office file was tampered to hide the malicious code from being detected by the Anti-Virus.

It work by comparing the number of code lines and characters on each lines with automated header removal between the decompiled p-code and the original source code. 

## Test run

1. Assuming Python 3 programming environment already configured by the user; execute `pip install oletools pcodedmp` to install the required dependencies.
2. Download the [pcode2code.py](https://github.com/Big5-sec/pcode2code/blob/master/pcode2code/pcode2code.py) source code and put it on the same directory as the [stomperDetector.py](https://github.com/hafiz-kamilin/vba_stomper_detector/blob/main/source/stomperDetector.py).
3. cd the console to the current directory and execute `python stomperDetector.py --dir ./<yourfile>.docm`.

Alternatively, if you prefer not to install python and follow these steps, you can download the compiled .exe version at [here](https://github.com/hafiz-kamilin/vba_stomper_detector/releases/tag/v1.0). To execute the file, cd the console to the current directory and execute `stomperDetector.exe --dir ./<yourfile>.docm`.

## Note

I have tested the stomperDetector on all stomped files provided by [clr2of8](https://github.com/clr2of8/VBAstomp) and this program is compatible with all files created by Microsoft Office program such as Excel and Word version 2003x32, 2007x32, 2010x32, 2013x32, 2013x64, 2016x32, 2016x64, 2019x32, 2019x64.

## Declaration

The `pcode2code.py` uploaded into this repository is taken directly from [Big5-sec](https://github.com/Big5-sec/pcode2code/blob/master/pcode2code/pcode2code.py) without any modification on 11 May 2021.
