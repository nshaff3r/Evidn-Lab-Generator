# Evidn Lab Generator

This script automates the creation of lab reports for Princeton's The Power is Ours campign.

## Steps for installing

1. Download the files in this repository. This can be done by navigating to [releases](https://github.com/nshaff3r/Evidn-Lab-Generator/releases) and downloading the .zip file containing everything.
2. Make sure that Python is installed. A great guide to install can be found [here](https://docs.python-guide.org/starting/installation/).
3. Unpack the downloaded zip of Evidn Lab Generator. 
4. Using a terminal or CLI on your computer (i.e Terminal on Mac, Powershell on Windows), navigate to the folder that was created by unpacking the zip. A guide for this can be found [here](https://medium.com/geekculture/basic-bash-commands-c54933183c89)
5. Install the necessary packages from requirements.txt. This can be done using the command
    
    ```pip install -r requirements.txt```

## Steps for running
1. Download the lab .csv files and put them in a folder. Title this folder the name of the lab building *(ex. Icahn)*. You should run this program seperately if you're dealing with multiple buildings.
    - **Make sure this folder is in the same location as the program files**
2. Open template.pptx and duplicate the slides for the amount of labs you have.
    - For example, if you wanted to make 4 lab reports, there should be 4 slides.
3. Run the program with the command

    ```python main.py```
4. Upon completion, the output file will appear, named "output.pptx"
