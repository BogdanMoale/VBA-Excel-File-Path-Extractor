# VBA Excel File Path Extractor

This VBA Excel script is designed to extract the paths of all files within a specified folder and its subfolders. It presents a user interface to select the folder, and then prompts for the type of files to search for using wildcard characters. The script then searches for matching files recursively in the selected folder and its subfolders, and outputs the file paths to an Excel sheet.

## Usage

1. Run the `Find_Files` subroutine.
2. Select the folder containing the files you want to search through.
3. Enter the file pattern using wildcard characters (e.g., `*.xls*`).
4. The script will populate the paths of matching files into the first column of the first sheet in the Excel workbook.
5. Run the `formulaLink` function to generate hyperlinks for each file path and its corresponding folder path.

## Functions

- `Find_Files`: Initiates the file search process.
- `formulaLink`: Generates hyperlinks for file paths and their corresponding folder paths, and populates them in the Excel sheet.

## Requirements

- Microsoft Excel

## Note

- This script is intended for use within Microsoft Excel and requires macro execution enabled.

