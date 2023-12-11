# Excel Worksheet Export

This VBA script is designed to automate the export of individual worksheets from an Excel workbook. It iterates through all worksheets except a specified dashboard worksheet, creates a copy of each worksheet, and saves it as a separate Excel file in a specified directory.

## Prerequisites

Make sure to enable the "Developer" tab in Excel to run this script. Follow these steps:

1. Click on "File" in the Excel ribbon.
2. Select "Options."
3. In the Excel Options dialog, choose "Customize Ribbon."
4. In the right pane, check the "Developer" option.
5. Click "OK" to apply the changes.

## Instructions

1. Open the Excel workbook containing the worksheets you want to export.
2. Press `Alt` + `F11` to open the Visual Basic for Applications (VBA) editor.
3. Insert a new module by right-clicking on any item in the Project Explorer, selecting "Insert," and then choosing "Module."
4. Copy and paste the provided VBA script into the module.
5. Close the VBA editor.
6. Press `Alt` + `F8` to open the "Macro" dialog.
7. Select "WorksheetExport" and click "Run."

## Configuration

- **wsDash:** Set the variable `wsDash` to the name of the dashboard worksheet that you want to exclude from the export.

- **filePathToSave:** Set the variable `filePathToSave` to the desired directory where the exported Excel files will be saved.
