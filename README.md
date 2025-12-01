# Property Data Processing Software

## Overview
This project is a Python application built for a municipal Assessor's Office to improve the efficiency of processing real estate sales data.
The tool loads MLS OneKey sales files, converts parcel numbers into print keys, pulls assessment information from RPS reports, formats the resulting dataset and generates annual analysis sheets including a bar chart comparing average vs. median sales prices.
All functionality is accessible through a Tkinter GUI with buttons.

## What The Project Does
* Reads MLS OneKey export files containing property sales from a specific timeframe.
* Converts parcel numbers into print_key identifiers used in RPS.
* Searches additional Excel files (reports generated via RPS) to pull:
  * Sales Price
  * 5217 Assessed Value (Value listed on the deed)
  * Current Assessed Value 
  * Condition Code
* Writes all retrieved data into new structured columns added to the sales file.
* Appends processed data into a hard-coded workbook (**annualsales.xlsx not included in repository due to confidential information**).
* Generates two annual summary sheets:
  1) A **summary sheet** with number of sales, average price, median price, AV/SP ratio, and more.
  2) A **data sheet** containing all appended yearly sales used for summary calculations.
* Creates a bar chart comparing average vs. median sales price based on user input and saves it as PNG.

## Why This Project Is Useful
Assessors often process hundreds of yearly sales to determine property value trends, equalization rates, and year-over-year changes.
This program automates tasks that are typically done manually, saving time and reducing errors by:
* Eliminating manual data lookup between RPS, MLS and additional websites/softwares
* Providing automatic formatting and structured output
* Producing annual statistics with one click
* Generating charts for reports and meetings

## How To Get Started 
### Requirements
* Python 3.x
* Required Libraries:
  * pandas
  * openpyxl
  * matplotlib
  * tkinter
  * pyinstaller (Optional unless .py is edited and needs to be repackaged as a new .exe)
* Excel files must follow the exact MLS OneKey and RPS formats used by the Assessor's Office

### Folder Setup
Place the following in the **same directory:**
  * MLS Sales Excel File
  * RPS Export Excel Files (Sales and Roll)
  * Program .exe file (and .py)
  * Annual Sales File (annualsales.xlsx **not included for confidentiality**; The program expects this file to exist for the Annual Sales Tools)

Important Note: The program will not run correctly if Excel files are placed outside of the directory that contains the .exe file. In addition, all excel files in use must be **closed** while using the program.

### Running The Program
Run the .exe file and a Tkinter GUI will appear with the following:
1) Data Processing with buttons:
   * Process Data
2) Annual Sales Tools with buttons:
   * Append Rows
   * Add New Sheets
   * Generate Chart

## Where Users Can Get Help
Because this project was built for a specific muncipal office with private datasets, outside users will need to adapt the file formats for their environment.
For help:
* Open an issue in this GitHub repository
* Contact the maintainer listed below

## Who Maintains This Project
This project is maintained by Madison Parente as part of work compared for a Town Assessor's Office initiative involving data automation and AI enchancements.

Contributions are welcome, but users must note:
* The tools relies on proprietary data formats (MLS OneKey + RPS)
* Confidential workbooks are excluded from this repository

  


