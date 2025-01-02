# Tracker Calendar Creator - Google Sheets Add-On
## Overview
Tracker Calendar Creator is a Google Sheets Add-On that simplifies the process of creating personalized tracker calendars. By inputting the year, your sheet ID, and weekly goal, the add-on automatically fills in a calendar with the proper structure, including weeks and daily averages and percentage of accomplished goal.

## Features
Input Form: User-friendly form to enter year, Sheet ID, weekly goal, and sheet name.
Dynamic Calendar: Automatically populates the sheet with a grid for each month, including week numbers and weekly counts.
Conditional Formatting: Adds conditional formatting for past dates and user-defined weekly goals.
Flexible Setup: Supports optional sheet name input, with default set to "Sheet1".

## Usage
The following link
[https://script.google.com/macros/s/AKfycbzCnrLVivI_MU_U5Ma9PvqWSErkBLwnG0H18HX8rjHU__0_0EACarBLjqcoH1OiFmJk/exec]
(https://script.google.com/macros/s/AKfycbzCnrLVivI_MU_U5Ma9PvqWSErkBLwnG0H18HX8rjHU__0_0EACarBLjqcoH1OiFmJk/exec)
opens the form that will automatically start popuating the sheet.
Fill out the form in the sidebar with:
Year: The year for which the calendar should be created.
Sheet ID: The Google Sheets document ID where the calendar will be created.
Weekly Goal: Number of days per week you aim to complete.
Sheet Name (optional): The specific sheet name within the document (default: Sheet1).
Submit the form, and the add-on will generate the calendar.
## Code Overview
The code consists of various functions to:

Generate grids for each month.
Add week numbers and calculate daily counts.
Implement conditional formatting for visualization.
Enable dynamic adjustment based on user input.
## Contributions
Feel free to fork the repository and contribute your improvements. Bugs and feature requests are welcomed as well!
