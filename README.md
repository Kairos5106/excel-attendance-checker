# excel-attendance-checker
A simple Python program to help map attendance between two Excel sheets

## Type of sheets
There will be two types of Excel sheet involved:
1. A **main** sheet: The Excel sheet that serves as a database for attendees (storing personal data and attendance across days), and
2. A **target** sheet: An Excel sheet that contains a list of attendees for that certain day (with timestamp)
> The target sheet is to be mapped onto the main sheet

# Dependencies
Make sure to install these dependencies to be able to use this program
1. Python v3.xx
2. Some Python packages:
- pandas
- openpyxl
```
pip install pandas openpyxl
```
> Run the command above in your command terminal to install the required Python packages 

# Instructions
1. Open your terminal and navigate to this directory. For example:
```
cd <your-directory-path>
```
> Example directory path: /Users/kevinhtw/Documents/coding-projects/claudia/excel-attendance-checker
2. Run the command below
```
py run.py
```
3. Follow the instructions on the terminal