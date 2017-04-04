This is a barebones script that automates Quartzy for our lab's workflow. 
The script supports multiple groups but the main function needs to be called for each group.

Setup:
This script requires selenium, chromedriver, libreoffice
The script assumes you are using Linux (tested on Ubuntu 16.04)
The script requires Libreoffice to be closed
The script works on python2.7. It should work on python3 but I haven't tested that

Operation:
Alter the filepaths, etc. as necessary in the indicated part of the code. Import the module and run it.

Sample Code:
import quartzy
quartzy.quartzy(0)

Workflow:
The script will generate requisition forms for each unique combination of vendor and grant number in addition to a summary CSV file.
The script will then concatenate all of these files into a single PDF file