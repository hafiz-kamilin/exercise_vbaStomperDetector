# VBA Stomper Detector

## Introduction

<p align = "center">
  <img src = "https://raw.githubusercontent.com/hafiz-kamilin/vba_stomper_detector/main/flowchart.png" width = "395" height = "653"/>
</p>

This is an example concept that shows how we can detect the disparity in compiled VBA macro (known as p-code) and VBA macro source code in DOCM file. This disparity reveal that the DOCM file was tampered to hide the malicious code from being detected by the Anti-Virus.

It work by implementing the number of code lines and characters comparison between the decompiled p-code and the original source code. 

## Test run

1. Assuming Python 3 programming environment already configured by the user; execute `pip install oletools pcodedmp` to install the required dependencies.
2. Download the [pcode2code.py](https://github.com/Big5-sec/pcode2code/blob/master/pcode2code/pcode2code.py) source code and put it on the same directory as the [stomperDetector.py](https://github.com/hafiz-kamilin/vba_stomper_detector/blob/main/stomperDetector.py).
3. cd the console to the current directory and execute `python stomperDetector.py --dir ./<yourfile>.docm`.

Alternatively, if you prefer not to install python and follow these steps, you can download the compiled .exe version at [here](https://github.com/hafiz-kamilin/vba_stomper_detector/releases/tag/v1.0). To execute the file, cd the console to the current directory and execute `stomperDetector.exe --dir ./<yourfile>.docm`.

## Note
I have coded this as a requirement for a job interview. I might not update the code again in future.

## Declaration

The pcode2code uploaded into this repository is taken directly from [Big5-sec](https://github.com/Big5-sec/pcode2code) without any modification on 11 May 2021.
