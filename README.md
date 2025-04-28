# pyreport

Purpose:
App for reading and concatenating multiple reports.

Setup:
pyreport file should be saved in a location (preferably a dedicated folder) containing 3 folders titled: "base", "extra", and "output".
base folder should contain the file you intend to build off of.
extra folder should contain files with information you would like to concatenate onto the base report.
output folder should be empty, this is where pyreport will place the finished excel file.

Useage:
1. Place main excel report into base folder
2. Place additional reports into extra folder
3. Ensure output folder is empty
4. Run pyreport script in terminal
5. pyreport will identify rows containing three anchor points per row
6. pyreport will then add remaining data from the extra reports onto the back end of the matching row in the base report
7. Outcome will be saved as a new excel file in the output folder
