#!/usr/bin/env python3
# -*- coding: utf-8 -*-

__author__ = "Mohd Hafizuddin Bin Kamilin"
__version__ = "1.0"
__date__ = "12 May 2021"

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

            print("\nThe file parsed is not a an Microsoft Office file.\n")

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

# extract the source code and p-code
def codeExtractor(vbaparser, argsInput):

    """ using the olevba tool """

    # extract the vba code from the file
    sourceCode = vbaparser.reveal()
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

        pCode=myfile.read()

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

    return refinedSourceCode, refinedpCode

# remove the unnecessary header from the code
def headerRemover(refinedSourceCode, refinedpCode):

    result = None
    # find the element that exist in both list and save in similarityFound list
    baseSet = set(refinedSourceCode)
    similarityFound = list(baseSet.intersection(refinedpCode))

    # if there is no similarity found, that mean the source code was totally destroyed (stomped)
    # if the only similarity found is the "End Sub", then there is none (stomped)
    if ((similarityFound != []) and (similarityFound != ["End Sub"])):

        # list format [element found, the position of the element in either refinedSourceCode or refinedpCode]
        elementValue = None
        elementPosition = None

        # find the position of each srcPosition's element in the refinedSourceCode and refinedpCode
        for element in range(len(similarityFound)):

            for line in range(len(refinedSourceCode)):

                # the first similar element found is stored in srcPosition
                if (refinedSourceCode[line] == similarityFound[element]):

                    if ((elementValue == None) and (elementPosition == None)):

                        elementValue = similarityFound[element]
                        elementPosition = line

                    elif (elementPosition >= line):

                        elementValue = similarityFound[element]
                        elementPosition = line

            for line in range(len(refinedpCode)):

                # the first similar element found is stored in srcPosition
                if (refinedpCode[line] == similarityFound[element]):

                    if (elementPosition >= line):

                        elementValue = similarityFound[element]
                        elementPosition = line

        # truncate every line in refinedSourceCode and stop when the line
        # contained the elementValue was found
        while True:

            if (refinedSourceCode[0] != elementValue):

                refinedSourceCode.pop(0)
            
            else:

                break

        # truncate every line in refinedpCode and stop when the line
        # contained the elementValue was found
        while True:

            if (refinedpCode[0] != elementValue):

                refinedpCode.pop(0)
            
            else:

                break

    else:

        result = True

    return result, refinedSourceCode, refinedpCode

# check if the source code is stomped or not
def checkStopedOrNot(refinedSourceCode, refinedpCode):

    """ check if the number of lines is same or not """

    if (len(refinedSourceCode) == len(refinedpCode)):

        for line in range(len(refinedSourceCode)):

            if (refinedSourceCode[line] != refinedpCode[line]):

                # source code is stomped
                return True

        # there is no disparity between the p-code and source code detected
        return False
    
    else:

        # source code is stomped
        return True

# main
if __name__ == "__main__":

    # parse the file
    vbaparser = parseFileCheck(args.dir)

    if (vbaparser is not None):

        # check if the vba macros exist or not
        existOrNot = checkVbaExist(vbaparser)

        # macros exist
        if (existOrNot is True):

            # extract the source code and p-code
            refinedSourceCode, refinedpCode = codeExtractor(vbaparser, args.dir)
            # remove header from the extracted codes
            result, refinedSourceCode, refinedpCode = headerRemover(refinedSourceCode, refinedpCode)

            if (result is None):

                # check if code is stomped or not
                result = checkStopedOrNot(refinedSourceCode, refinedpCode)
        
            if (result is True):

                print("\nThe length of decompiled p-code and source code is not same.")
                print("This shows the file contains VBA stomped code!\n")
            
            else:

                print("\nThe length of decompiled p-code and source code is same.")
                print("Although this shows the file does not contains VBA stomped code, user should still proceed with caution.\n")

        else:

            print("\nFile check end here.\n")
