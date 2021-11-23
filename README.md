# MRRO
MALTA REPROGRAPHIC RIGHTS ORGANISATION

**MRRO.py** - Reads in separate publisher submissions in the form of spreadsheets (Excel)
following the default template. Works out payout per book based on given formula.
Exports three CSV files:

- all books from all publishers including all relevant parameters and calculations
- publisher payments
- author payments

All spreadsheets should be in a folder with nothing else in the folder (except files
created by this software itself which are added in the same folder and can remain there)

### Usage:
When you run the application you'll have two prompts:

- First prompt - enter path of folder with spreadsheets (option-command-c will copy
the path on macOS when folder is selected in Finder)
- Second prompt - enter amount of money to be distributed (total funds - admin charges)

Three files will be created each time the software runs with the required outputs.
The files will be in CSV format, and therefore, can be opened with Excel
or any other spreadsheet software. Files have a timestamp in the file name showing when they
were created.

Application runs on macOS only. No Windows version is available at present

### Notes:

- Authors must be seperated by commas, for example `John Smith, Mary White`
- Publisher name should follow the colon in cell `"Name of Company/Self Published Author: "`
- Rows with entries `"List of books published in YYYY"` are required and should not be modified in anyway, except for changing the year
- Retail price should not include currency, so input `5.25` for example
- The column `"Melitensia (1), Adult (2), Children's (3)"` should only include the numbers `1`, `2`, or `3`
- No other changes should be make to the template spreadsheet



### Known Issues

- Currently, this code is Mac only. However, the only platform-specific aspect of the code is the macOS-style path names which use /. Therefore, code should run fine on Unix/Linux systems. The forward slash can be changed to a backslash to run on Windows. The code should then run fine on Windows, too. However, this hasn't been tested in this version
- The template must be followed and not altered in any way for the code to run smoothly. In future versions, the template is likely to be changed to make data entry and collection more robust
