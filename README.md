# DuplicateRemover
Basic python tool for removing duplicate entries in very large excel dataset of blood test results.

# How to use

Have the excel file to be filtered, the python script and the config.ini file in the same folder.

Open a command prompt or powershell in the same folder and enter "python filter.py" without any other arguments

Input file and sheet index to be sorted are entered in the config.ini file. This is also where you define what name you want the output file to have. CAREFUL, if a file already exists with the name given for the output file, it will be deleted/overwritten without any warnings.

This script relies on the first column being patient id's, and on patient id's being sorted in order such that all instances of the same patient id are consecutive.

All required packages are available through pip if needed, not sure if any are non native.
If you can't get your python environment to work I can try compiling to a .exe, but not sure how large it would be.

-adam 05-09-2018

# Changelog

Fixed bug relating to date parsing sorting days before months, added function call to strptime with d/m/y format. If this fails will require customisation of date parsing in config.

-adam 07-09-2018

Cleaned up folder for github.

Created small example file with personal information redacted.

-adam 19-04-2019
