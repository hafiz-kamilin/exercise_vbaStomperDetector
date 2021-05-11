#!/usr/bin/env python3
# -*- coding: utf-8 -*-

__author__ = "Mohd Hafizuddin Bin Kamilin"
__version__ = "0.0.1"
__date__ = "11 May 2021"

# for parsing the vba source file
from oletools.olevba import VBA_Parser, TYPE_OLE, TYPE_OpenXML, TYPE_Word2003_XML, TYPE_MHTML
# for parsing the compiled source file
import pcode2code
# parsing the file location from cli
import argparse

# initialize argparse
parser = argparse.ArgumentParser()
# define dir as argument to pass the file location
parser.add_argument("--dir")
# string is stored as args
args = parser.parse_args()

# check if the parsed file is valid or not
def parseFileCheck(argsInput):

    # args is empty; cannot proceed
    if (argsInput is None):

        print("\nCommand line argument is not complete!\n")
        print("Please run the program once again by following either one of these examples:")
        print("1. python stomperDetector.py --dir ./Question_1.docm")
        print("2. python stomperDetector.py --dir 'C:/Users/<Profile>/answers/detection/Question_1.docm'\n")

    # not empty; can proceed
    else:

        try:

            vbaparser = VBA_Parser(argsInput)

        # file does not exist
        except FileNotFoundError:

            print("\nFile does not exist!")
            print("Please run the program once again by following either one of these examples:")
            print("1. python stomperDetector.py --dir ./Question_1.docm")
            print("2. python stomperDetector.py --dir 'C:/Users/<Profile>/answers/detection/Question_1.docm'\n")

        # most likely a wrong file was parsed
        except:

            print("\nThe file parsed is not a DOCM file.\n")

        # if the file can be opened...
        else:

            return vbaparser

# check if vba macros exist inside the file or not
def checkVbaExist(vbaparser):
    
    if vbaparser.detect_vba_macros():

        print("\nVBA Macros found.")
        return True

    else:

        print("\nNo VBA Macros found.")
        return False

def compareForAbnormality(vbaparser, argsInput):

    """ using the olevba tool """

    # extract the vba code from the file
    for (_, _, _, sourceCode) in vbaparser.extract_macros():

        # remove the code attributes from the extracted source code
        # NOTE: below shows the example of the attributes from the sourceCode,
        #   it is useless for comparing it with decoded p-code.
        #   Attribute VB_Name = "ThisDocument"        
        #   Attribute VB_Base = "1Normal.ThisDocument"
        #   Attribute VB_GlobalNameSpace = False      
        #   Attribute VB_Creatable = False
        #   Attribute VB_PredeclaredId = True
        #   Attribute VB_Exposed = True
        #   Attribute VB_TemplateDerived = True
        #   Attribute VB_Customizable = True
        sourceCode = sourceCode.split("\n",8)[8]
        # convert \r\n to \n
        sourceCode = "\n".join(sourceCode.splitlines())
        # convert str to list where \n act as a seperator
        sourceCode = list(sourceCode.split("\n"))

        # remove trailing and leading space from list's str
        for line in range(len(sourceCode)):

            sourceCode[line] = sourceCode[line].lstrip()

        refinedSourceCode = []

        # remove empty str from list
        for line in range(len(sourceCode)):

            if (sourceCode[line] != ""):

                refinedSourceCode.append(sourceCode[line])

    """ using the pcode2code """

    # use external python code to decode the p code
    pcode2code.process(argsInput, "decoded_pcode.txt")

    # read the decoded text file 
    with open ("decoded_pcode.txt", "r") as myfile:

        decodedSource=myfile.read()

    # strip the decoded file information
    # NOTE: below shows the example of decoded file information that need to be stripped
    #    stream : VBA/ThisDocument - 74307 bytes
    #
    #    ########################################
    #
    #
    pCode = decodedSource.split("\n",6)[6]
    # convert str to list where \n act as a seperator
    pCode = list(pCode.split("\n"))

    # remove trailing and leading space from list's str
    for line in range(len(pCode)):

        pCode[line] = pCode[line].lstrip()

    refinedpCode = []

    # remove empty str from list
    for line in range(len(pCode)):

        if (pCode[line] != ""):

            refinedpCode.append(pCode[line])

    """ check if the number of lines is same or not """

    if (len(refinedSourceCode) == (refinedpCode)):

        # the source code is not stomped
        return False
    
    else:

        # source code is stomped
        return True

# main
if __name__ == "__main__":

    # parse the docm file
    vbaparser = parseFileCheck(args.dir)

    if (vbaparser is not None):

        # check if the vba macros exist or not
        existOrNot = checkVbaExist(vbaparser)

        # macros exist
        if (existOrNot is True):

            # check if code is stomped or not
            result = compareForAbnormality(vbaparser, args.dir)
        
            if (result is True):

                print("\nThe length of decompiled p-code and source code is not same.")
                print("This shows the DOCM file contains VBA stomped code!\n")
            
            else:

                print("\nThe length of decompiled p-code and source code is same.")
                print("Although this shows the DOCM file does not contains VBA stomped code, user should still proceed with caution.\n")

        else:

            print("\nFile check end here.\n")
