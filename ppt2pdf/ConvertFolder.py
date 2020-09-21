import sys
import os
import comtypes.client
import click
from PyInquirer import prompt, Separator
from pyfiglet import Figlet


def createCheckBoxes(files):
    modified_list = []
    for file in files:
        temp = {}
        temp['name'] = file
        modified_list.append(temp)
    return modified_list

@click.command(name="ppt2pdf")
@click.argument("input_folder", required=True,type=click.Path(exists=True, resolve_path=True), nargs=1)
@click.option("-o","--output","output",type=click.Path(dir_okay=True, resolve_path=True), nargs=1, help="Output Folder default current folder")
@click.option("-s","--select",is_flag=True, help="Recursively Select Files")
def ConvertHere(input_folder,output,select):
    #%% Add final slash at end
    input_folder += "\\"
    if(not output):
        output=input_folder
    else:
        output += "\\"
    #%% Get files in input folder
    input_file_paths = os.listdir(input_folder)
    filtered_files = []

    for file in input_file_paths:
        # Skip if file does not contain a power point extension
        if  file.lower().endswith((".ppt", ".pptx")):
            filtered_files.append(file)

    if(select):
        user_choices = [
            {
                'type': 'checkbox',
                'qmark': '?',
                'message': 'Select the files you want to Convert',
                'name': 'convert',
                'choices':createCheckBoxes(filtered_files)
            }
        ]
        selected_files = prompt(user_choices)
        convert_files = selected_files['convert']
    else:
        convert_files = filtered_files
    #%% Convert each file
    with click.progressbar(convert_files,label="Status", show_pos=True) as bar:
        for input_file_name in bar:
            # # Create input file path
            input_file_path = os.path.join(input_folder, input_file_name)
            # # Create powerpoint application object
            powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
            # # Set visibility to minimize
            powerpoint.Visible = 1
            # # Open the powerpoint slides
            slides = powerpoint.Presentations.Open(input_file_path)
            # # Get base file name
            file_name = os.path.splitext(input_file_name)[0]
            # # Create output file path
            output_file_path = os.path.join(output, file_name + ".pdf")
            # # Save as PDF (formatType = 32)
            slides.SaveAs(output_file_path, 32)
            # # Close the slide deck
            slides.Close()

if __name__ == '__main__':
    f = Figlet(font='cybermedium')
    click.echo(click.style(f.renderText('PPT => PDF'), fg="red"))
    ConvertHere()