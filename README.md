# Excel-VBA-Macros

## Overview
This repository contains a collection of VBA macros designed to automate various tasks in Excel. The scripts provided here can help streamline repetitive processes, improve efficiency, and enable more effective data management.

## Contents
- **Workbook Opener Macro:** A script to open multiple workbooks automatically upon opening the main workbook.

## Usage
1. Download or clone the repository to your local machine.
2. Open the Excel workbook where you want to add the macro.
3. Press `Alt + F11` to open the VBA editor.
4. Copy the desired script from the repository and paste it into the VBA editor.
5. Customize the file paths and any other parameters as needed.
6. Save the workbook as a macro-enabled file (`.xlsm`).

## Example Code
```vba
Private Sub Workbook_open()
    Dim wb1 As Workbook
    Dim myfilename1 As String
    Dim wb2 As Workbook
    Dim myfilename2 As String
    myfilename1 = "FILEPATH"
    myfilename2 = "FILEPATH"
    '~~> Open the workbook and pass it to workbook object variable
    Set wb1 = Workbooks.Open(myfilename1)
    Set wb2 = Workbooks.Open(myfilename2)
End Sub
