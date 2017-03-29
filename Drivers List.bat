REM batch script takes directory data and coverts it to a text file on the desktop called driver_list
REM created for HR to export from Box file with nested folders for each eligible driver, to create list to send monthly of who can drive for the company
DIR %1 /O > C:\Users\%USERNAME%\Desktop\drivers_list.txt