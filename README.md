# Batch Update Tool

**Batch Update Tool** is an Excel VBA add-in that provides tools for automatically updating the name of a range of sheets or cell values with a specified sequence. You can also save the settings as "Quick Actions" and execute them quickly by selecting them in the "Quick Actions" menu.

## Key Features

**Batch Sheet Name Update**

![Batch Sheet Name Update](README.assets/Batch%20Sheet%20Name%20Update.png)

- Update the name of a set of sheets with a specified sequence
- Customize the sequence by modifying different parameters
- Update all the sheets in the Workbook or select the range of sheets you would like to update

**Batch Cell Update**

![Batch Cell Update](README.assets/Batch%20Cell%20Update.png)

- Update the value of a set of cells with a specified sequence
- Customize the sequence by modifying different parameters
- Add your custom sequence, and the program will iterate through the sequence in a loop
- Select a range of sheets that you would like to perform the cell update

**Advanced Sheet Copier**

![Advanced Sheet Copier](README.assets/Advanced%20Sheet%20Copier.png)

- Copy a set of sheets a specified number of times
- Select the position where the sheets are pasted

## Installation

To install this Add-in to your computer:

1. Download  [Batch Update Tool.xlam](Batch%20Update%20Tool.xlam) to your computer
2. Enable this Add-in in Microsoft Excel
3. Add the Macros to your Quick Access Toolbar to access it

You can refer to the [Installation Guide](Documentation/Installation%20Guide.md) for a step-by-step tutorial on installing this Add-in to your computer.

## Issues?

Though there are no known bugs so far, the exception handling is not implemented in this program comprehensively, so a wrong input can cause the program to crash. Please follow the [User Guide](Documentation/User%20Guide.md) to get started with the program and learn how to operate the program correctly.

For any issues regarding the program, please don't hesitate and write an email to [benchoy352@gmail.com](mailto:benchoy352@gmail.com).

## Acknowledgment

Implementation of selecting a folder in VBA is adopted from [TheSpreadsheetGuru](https://www.thespreadsheetguru.com/the-code-vault/vba-code-to-select-folder-path).