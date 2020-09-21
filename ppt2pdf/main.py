import sys
import os
import comtypes.client
from pyfiglet import Figlet

class PPT2PDF():

    def __init__(self):
        f = Figlet(font='isometric1')
        print(f.renderText('PPT 2 PDF'))
        


    def _generateOutputFilename(self,outputFilename):
        output = os.path.splitext(outputFilename);
        output=os.path.abspath(output[0]+".pdf");
        return output;


    def convert(self, inputFilePath,outputFilePath):
        
        print("Your Input file is at:")
        print(inputFilePath)

        if(not outputFilePath):
            outputFilePath = self._generateOutputFilename(inputFilePath);
        
        print("Your Output file will be at:")
        print(outputFilePath);

        # %% Create powerpoint application object
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        #%% Set visibility to minimize
        powerpoint.Visible = 1
        #%% Open the powerpoint slides
        slides = powerpoint.Presentations.Open(inputFilePath)
        #%% Save as PDF (formatType = 32)
        slides.SaveAs(outputFilePath, 32)
        #%% Close the slide deck
        slides.Close()


    def convertSingleFile(self,inputFilePath,output):
        if(not output):
            output=self._generateOutputFilename(inputFilePath)
        self.convert(inputFilePath,output)

    def testImport(self):
        print("Yesssssssssssssssssssssssssssss");