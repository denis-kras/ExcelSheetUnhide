# Passing Parameters to the script
[CmdletBinding()]
param(
    [Parameter()][string]$in,
    [Parameter()][string]$out,
    [Parameter()][switch]$c,
    [Parameter()][switch]$u,
    [Parameter()][switch]$e,
    [Parameter()][switch]$f,
    [Parameter()][switch]$v,
    [Parameter()][switch]$l,
    [Parameter()][switch]$h
)

#==================================

# Basic Variables
$Version = "v1.0.1"
$Name = "ExcelSheetUnhide"
$Author = "DenK"
$SaveFile = $false

#==================================

# Functions

# Checking Path with some error handling of Test-Path
Function CheckPath($Path) {
    $Existing = $true
    
    Try {
        $Existing = Test-Path $Path
    }
    Catch {
        $Existing = $false
    }

    Return $Existing
}

# Execute Excel editor and open the file
Function OpenExcelFile($ExcelInputFile, [boolean]$ApplicationVisability) {
    Try {
        # Execute Excel editor through ComObject, If there is no Excel Editor installed, throw an error
        $Global:Excel1 = New-Object -ComObject Excel.Application -ErrorAction Stop
    }
    Catch {
        Write-Host "Excel ComObject Couldn't be executed. Check If Excel is installed." -ForegroundColor Red
        Exit
    }
    
    # By default the ".Visible" Property is set to $false and Excel editor application will not be visible
    If ($ApplicationVisability) {
        $Excel1.Visible = $true
    } Else {
        $Excel1.Visible = $false
    }

    # 3 is Disable all the Macros upon opening - no alerts shown in Excel editor.
    # 1 is enable, 2 is using the Security tab settings.
    $Excel1.AutomationSecurity = 3

    # For reference only:
    # As part of Workbooks.Open method, there is "Editable" object, which is omitted and $false by default.
    # This object relates to Excel 4.0 addon, Which will be "disabled" in turms of editing.
    # Which is a good thing.
    Try {
        # Open the Excel file
        $Global:Workbooks = $Excel1.Workbooks.Open($ExcelInputFile)
    }
    Catch {
        Write-Host "Couldn't Open the Excel file." -ForegroundColor Red
        Exit
    }
    # Microsoft suggests using it twice, before and after the programmatically opening
    # of the file to avoid malicious subversion.
    $Excel1.AutomationSecurity = 3
    
    # Define $Sheets Script wide
    $Global:Sheets = $Workbooks.sheets

}

# Close the excel file and excel editor
Function CloseExcelFile($ExcelOutputFile, [boolean]$SaveFileVariable) {
    $GotoExit = $false
    
    # If there is no need to "Save As" the file, just close the Excel editor
    If (!$SaveFileVariable) {
        # $alse will omit errors and alerts
        $Global:Workbooks.Close($false)
    # If you need to "Save As" the file
    } Else {
        # Take only the path to the output file
        $OutputPath = Split-Path $ExcelOutputFile
        # Anc check if the path is existing before saving the file there
        $ExistingOutputPath = CheckPath($OutputPath)

        If (!$ExistingOutputPath) {
            Write-Host "Output Path '"$OutputPath"' is Non-Existent." -ForegroundColor Red
        } Else {
            If (CheckPath($ExcelOutputFile)) {
                Try {
                    # Since the ".SaveAs" method can't overwrite the file silently, it needs to be deleted first
                    Remove-Item -Path $ExcelOutputFile -Force -ErrorAction Stop
                }
                Catch {
                    Write-Host "Couldn't overwrite the file <"$ExcelOutputFile">" -ForegroundColor Red
                    # Excel editor needs to be closed before terminating the script. This variable helps do just that, since "GoTo" isn't considered best practice
                    $GotoExit = $true
                }
            }

            If (!$GotoExit) {
                Try {
                    $Workbooks.SaveAs($ExcelOutputFile)
                }
                Catch {
                    Write-Host "Something went wrong during file export, check the <-out> switch!" -ForegroundColor Red
                }
            }    
        }
        $Global:Workbooks.Close($false)
    }
    
    $Global:Excel1.Quit()

    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Global:Excel1) | Out-Null

    # After closing the Excel editor, check if the output file was really written
    If (CheckPath($ExcelOutputFile)) {
        Write-Host "The file saved" -ForegroundColor Green
    } Else {
        Write-Host "Error Saving <"$ExcelOutputFile" >" -ForegroundColor Red
    }
}

# Show Visible property of all the Sheets inside Excel document
Function CheckSheets {
    Write-Host "Check for hidden Excel sheets"
    ForEach ($Sheet in $Sheets) {
        Switch ($Sheet.Visible) {
            -1 { Write-Host $Sheet.Name "| Visible | <"$Sheet.Visible">" }
            2 { Write-Host $Sheet.Name "| Very Hidden | <"$Sheet.Visible">" }
            default { Write-Host $Sheet.Name "| <"$Sheet.Visible">" }
        }
    }
    BreakLine
}

# Function to unhide only Sheets that aren't valued "-1"
Function UnhideSheets {
    $HiddenSpreadShets = $false
    ForEach ($Sheet in $Sheets) {
        If ($Sheet.Visible -ne -1) {
            Write-Host "Changing Spreadsheet <"$Sheet.Name"> Visible value of <"$Sheet.Visible">"
            $Sheet.Visible = $true

            # Check if the change of Visible property of current Sheet really succeeded
            If ($Sheet.Visible -eq -1) {
                Write-Host "Success"
                $HiddenSpreadShets = $true
            } Else {
                Write-Host "No Success"
            }
        }
    }
    
    If (!$HiddenSpreadShets) {
        Write-Host "There are no Hidden Spreadsheets."
    }

    BreakLine

    # Show all the Sheets again after the change
    CheckSheets
}

Function BreakLine {
    Write-Host "========================"
}
#=======================================

# Main

Write-Host $Name $Version
Write-Host "Author: $Author"
Write-Host "Note that Excel editor is opened in the background with Macros disabled. Better use the script in VM."
Write-Host "-h: Show help, Syntax and Examples. Other switches will be omitted."
Write-Host "The Script is licensed under GNU General Public License v3.0"
Write-Host "Visit Script's site page with manual and explanations: https://www.optimizationcore.com/excelsheetunhide"
Write-Host "Visit GitHub page: https://github.com/denk-core/ExcelSheetUnhide"
BreakLine

If (!$in -and !$out -and !$c -and !$u -and !$e -and !$f -and !$v -and !$l -and !$h) {
    $h = $true
}

# If -h switch is included
If ($h) {
    Write-Host "Help, Syntax, Examples" -ForegroundColor Blue
    BreakLine
    Write-Host $Name $Version
    Write-Host "Author: $Author"
    Write-Host "-in: Input Excel file. Must include full path."
    Write-Host "-out: Output Excel file that you want to save as after editing. Must include full path."
    Write-Host "-c: Check for Hidden Sheets only."
    Write-Host "-u: Check if Sheets can be Unhidden. You can use it with <-c> to check the Before Uhide status."
    Write-Host "-e: Must be used with <-u> and <-out>. Export unhidden file and save as <-out>."
    Write-Host "-f: Force overwrite the existing file in <-out>."
    Write-Host "-v: Make the Excel execution visible. Since the Application that opens the Excel file opens in the background, you can show it with this switch."
    Write-Host "-l: Leave the Excel opened and finish the switch without doing anything. Works only with <-v> switch. <-e>, <-out>, <-f> will be omitted."
    Write-Host "---------------------------"
    Write-Host "Usage Example 1 - Check only for the hidden Sheets inside Excel document:"
    Write-Host ".\ExcelUnhideSHeet.ps1 -c -in 'C:\YourMalicious.xls'" -ForegroundColor Blue
    Write-Host "---------------------------"
    Write-Host "Usage Example 2 - Open Excel Application, unhide the Sheets and terminate the script, leaving you the option of editing it or saving as you want:"
    Write-Host ".\ExcelUnhideSHeet.ps1 -u -v -l -in 'C:\YourMalicious.xls'" -ForegroundColor Blue
    Write-Host "---------------------------"
    Write-Host "Usage Example 3 - Don't show Excel Application, Shows Sheets before Unhide, Unhide, Save as and overwrite if output file existent:"
    Write-Host ".\ExcelUnhideSHeet.ps1 -c -u -e -f -in 'C:\YourMalicious.xls' -out 'C:\YourUnhidden.xls'" -ForegroundColor Blue
    Exit
}

If ($l -and (!$v)) {
    Write-Host "<-l> switch can be used only with <-v>. There is no point Leaving Excel opened if it is hidden." -ForegroundColor Red
    Exit
}

If ($e) {
    If (!$u) {
        Write-Host "Switch <-u> wasn't specified. Nothing to export." -ForegroundColor Red
        Exit
    }

    If (!$out) {
        Write-Host "Switch <-out> wasn't specified. Nothing to export." -ForegroundColor Red
        Exit
    }

    If ($u -and $out) {
        $SaveFile = $true
    }
}

if ($out) {
    $ExistingExcelOutputFile = CheckPath($out)
    If ($ExistingExcelOutputFile -and !$f) {
        Write-Host "The file <"$out"> is existent. Use <-f> switch to overwrite." -ForegroundColor Red
        Exit
    }
}

$ExistingExcelInputFile = CheckPath($in)
If (!$ExistingExcelInputFile) {
    Write-Host "Input file Non-existent. Exiting" -ForegroundColor Red
    Exit
}

# Open the Excel file
OpenExcelFile -ExcelInputFile $in -ApplicationVisability $v

If ($c) {
    CheckSheets
}

If ($u) {
    UnhideSheets
}

If ($l) {
    Write-Host "<-l> was used, Terminating script - Leaving Excel Opened."
    Exit
}

CloseExcelFile -ExcelOutputFile $out -SaveFileVariable $SaveFile