<h1 align="center">ExcelSheetUnhide</h1>


<!-- ABOUT THE PROJECT -->
## About The Project

The main purpose of “ExcelSheetUnhide” is to Unhide Hidden Excel Spread Sheets of the Excel files. Something that went popular on malicious Excel files that use Excel 4 Macros. Since the adversaries use “very hidden” value of the Excel Workbook, there is no way by un-hiding the workbook from the GUI of Microsoft Excel. You need to put some effort of doing it. For this particular reason I wrote the “ExcelSheetUnhide” in Powershell to automate this task and make it easier.

ExcelSheetUnhide – Unhide Hidden Sheets, to use this tool, you will need an Excel editor installed that can be automated through the “Excel.Application” ComObject (Microsoft Excel is the default option). Since, the Excel file is opened with default Excel editor as an object in order to numerate the Sheets / Workbooks and unhide them. As you can understand, this script can’t be considered static, since it actually executes the Excel file. This is how the ComObject of Application works. Though, I’m using the “AutomationSecurity” Application property. This property has “MsoAutomationSecurity” enumeration with value of 3 to disable all the macros twice (before and after the file opening) as Microsoft suggests. This method won’t show any alerts if you choose the “-v” Visible Excel option, but I checked a “Workbook_Open” VBA, and it didn't execute. I still suggest a VM environment while using this script.

Microsoft makes it hard using the “Workbook.Open” method in Powershell for no reason. Since this method was initially used in VBA scripts, it is much easier there. The problem is “Editable” object of this method, which is omitted in this script. This object is set to “$false” by default if omitted, thus “disabling” the Excel 4 Macro editing, which is a good thing until Microsoft decides to change the default value. I don’t believe it will happen though, since Excel 4.0 Macro is outdated, but still prefer to use best practices of declaring important objects (security wise).



<!-- GETTING STARTED -->
## Getting Started

### Prerequisites

* Powershell 5.1 or higher.
* Microsoft Excel 2010 or higher (Didn't test lower versions).

### Installation

1. Download the “ExcelSheetUnhide.ps1” script from my ExcelSheetUnhide GitHub page.
2. Run “powershell.exe”
3. Navigate to the folder that the “ExcelSheetUnhide.ps1” was extracted:
    ```cmd
    cd C:\Users\User1\Downloads\ExcelSheetUnhide
    ```

4. Execute the script to check the help notes:
    ```cmd
    .\ExcelSheetUnhide.ps1
    ```

You can also use the “-h” switch:
    ```cmd
    .\ExcelSheetUnhide.ps1 -h
    ```

The help notes have all the information you need about all available switches and three usage case scenarios.



<!-- USAGE EXAMPLES -->
## Usage

### Usage Example 1 – Check

Before making some changes or un-hiding Excel file Sheets / Workbooks the one would want to Check if any Workbooks are really hidden. Syntax:
```cmd
.\ExcelSheetUnhide.ps1 -c -in 'C:\YourMalicious.xls'
```

“-c” is the “Check” option and “-in” is the full path to the input Excel file. The output shows all the Sheets / Workbooks that are in the Excel file from the input and shows all the “Visibility” states of each Sheet. Since it uses the output from the Powershell object currently my interest was of values of “-1”, which is regular visible Sheet and “2”, which is the value of “very hidden” Sheet / Workbook. Currently, these are the two values that will be shown as strings and other values than “-1” or “2”, will be shown as numeric values. Which probably will also be hidden, but can be unhidden in more standard ways from the GUI. Off course ExcelSheetUnhide can reveal them also.

### Usage Example 2 – Unhide Hidden Excel Sheets

After you saw with “Usage Example 1” which Sheets / Workbooks are available and which are hidden – you would want to Unhide them. Syntax:
```cmd
.\ExcelSheetUnhide.ps1 -u -v -l -in 'C:\YourMalicious.xls'
```

“-u” is “Un-hiding” the Sheets, “-v” is Making the Excel application “Visible”, “-l” is terminating the script “Leaving” the Excel application opened.

Since the Excel editor application is already opening the file it is easier to manipulate the file inside of it after the Sheets / Workbooks were revealed. You can save it as another file, without changing the input file and edit it however you like. The Excel application is opened in “AutomationSecurity” mode “3” with disabled macros without any notice messages. So, the macros won’t be executed on file opening. You can save your file with revealed Sheets / Workbooks, close this instance of Excel application and reopen it again with default Excel application settings with a message to [Enable Content] (that is if you didn't change the default security settings).

### Usage Example 3 – Console only usage

If you don’t like to mess with the GUI of Excel application, you can use fully command scripted environment. Just remember that Excel application will be opened in the background anyway. Of course using the same “AutomationSecurity” mode “3”. This is Script wide and is not changing anywhere and is not affected by any of the switches. Syntax:
```cmd
.\ExcelSheetUnhide.ps1 -c -u -e -f -in 'C:\YourMalicious.xls' -out 'C:\YourUnhidden.xls'
```

“-c” is the “Check” from the first example, “-u” is the same “Unhide” hidden Excel sheets from the second example, “-e” is “Exporting” the Unhidden changes that were made by the “-u” Unhidden change to the Excel document or “Saving As”, “-f” is “Forcing” overwrite of the existing Output file, “-out” is off course the destination file to Export the changes to.

The logic is simple, probably you would want to see the Sheets and their Visibility values before the change and after the change, off course exporting and saving the Unhidden change to the Output file and Overwriting it if necessary.



<!-- LICENSE -->
## License

Distributed under the MIT License. See `LICENSE.txt` for more information.



# History

1.0.1 - 06.04.2020
* Fix: Help notes

1.0.0 - 05.04.2020
* Initial release