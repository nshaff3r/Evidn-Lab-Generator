# Evidn-Lab-Generator

This script automates the creation of lab reports for Princeton's the Power is Ours campign. If there are any bugs, please add an issue and describe the problem.

## Steps for installing

1. Download the files in this repository. This can be done by navigating to [releases](https://github.com/nshaff3r/Evidn-Lab-Generator/releases) and downloading the .zip file containing everything.
2. Make sure that Python is installed. The latest version can be found [here](https://www.python.org/downloads/).
3. Unpack the downloaded zip. 
4. Using a terminal or CLI on your computer, navigate to the folder that was created by unpacking the zip 
    - You can do this on Mac by right clicking on the folder name, going to the services tab, and clicking "New Terminal at Folder"
    - You can do this on Windows by opening the folder, right clicking anywhere in the folder while holding the shift key, and clicking "Open PowerShell window here"
    - If you're using Linux, you probably know how to do this already ;)
5. Install the necessary packages from requirements.txt. This can be done using the command
    
    ```pip install -r requirements.txt```

## Steps for running
1. Download the lab .csv files and put them in a folder. Title this folder the name of the lab building *(ex. Icahn)*. You should run this program seperately if you're dealing with multiple buildings.
    - **Make sure this folder is in the same location as the program files**
2. Open template.pptx and duplicate the slides for the amount of labs you have.
    - For example, if you wanted to make 4 lab reports, there should be 4 slides.
3. Run the program with the command

    ```python labs.py```
4. Upon completion, the output file will appear, named "output.pptx"
