import sys
import os
import comtypes.client
import click
from pyfiglet import Figlet

@click.command(name="ppt2pdf")
@click.argument("input_file", required=True,type=click.Path(exists=True, resolve_path=True),nargs=1)
@click.option("-o","--output", "output", type=click.Path(resolve_path=True))
def convertPPT2PDF(input_file,output):
    print("Your Input file is at:")
    click.echo(input_file)
    print("Your Output file will be at:")
    if(not output):
        output = os.path.splitext(input_file);
        output=os.path.abspath(output[0]+".pdf");
    print(output);
    
    # %% Create powerpoint application object
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    #%% Set visibility to minimize
    powerpoint.Visible = 1
    #%% Open the powerpoint slides
    slides = powerpoint.Presentations.Open(input_file)
    #%% Save as PDF (formatType = 32)
    slides.SaveAs(output, 32)
    #%% Close the slide deck
    slides.Close()

if __name__ == '__main__':
    f = Figlet(font='isometric1')
    print(f.renderText('PPT 2 PDF'))
    convertPPT2PDF()