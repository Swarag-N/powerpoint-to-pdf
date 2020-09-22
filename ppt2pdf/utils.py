import sys
import os

def createCheckBoxes(files):
    modified_list = []
    for file in files:
        temp = {}
        temp['name'] = file
        modified_list.append(temp)
    return modified_list

def generateOutputFilename(outputFilename):
    output = os.path.splitext(outputFilename);
    output=os.path.abspath(output[0]+".pdf");
    return output;
