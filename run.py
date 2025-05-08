import pandas as pd
import os
import shutil

from helpers import (
    add_new_names_to_sheet,
    clean_data,
    create_date_col_in_main,
    create_reverse_header_map,
    mark_absentees,
    mark_attendees,
    preview_sheet_data,
    confirm_file_selection,
    safe_listdir,
    sort_names,
)

# Config
name_column = "Nama Penuh (Nitih IC)".strip() # name columns must be the same in both main and target sheets
header_map = {
    "Nama Penuh (Nitih IC)": "name",
    "Umur": "age",
    "No. HP\n*Enti nadai hp, engkah no. hp apai/indai\nChunto: 012-3456789 (Apai)": "phone_number",
}
reversed_header_map = create_reverse_header_map(header_map)

# For developers only
debug = False

# Folder paths
main_relative_path = "/main"
processed_relative_path = "/processed"
target_relative_path = "/target"
current_path = os.getcwd()

# Construct absolute paths
main_path = current_path + main_relative_path
processed_path = current_path + processed_relative_path
target_path = current_path + target_relative_path

print(f"\nYour current path is {current_path}\n")

# Get file lists with error handling
processed_files = safe_listdir(processed_path, "processed", verbose=False)

## Select target file first before main
target_files = safe_listdir(target_path, "target", verbose=True)
confirm_target_file = confirm_file_selection(
  files=target_files, 
  type="target"
)

## Select main file
### TODO: Automatically select the most recent main file based on date or let user select manually
main_files = safe_listdir(main_path, "main", verbose=True)
### When first time entry, create new file from selected target and end program
confirm_main_file = ""
first_time_use = len(main_files) == 0
if first_time_use:
  #### Create a copy of target as main sheet
  confirm_main_file = confirm_target_file
else:
  confirm_main_file = confirm_file_selection(
    files=main_files, 
    type="main"
  )

# Read target sheet
target_file_path = target_path + "/" + confirm_target_file
if debug: print(f"\nTarget file is at {target_file_path}")

target_data = pd.read_excel(target_file_path)
if debug: preview_sheet_data(target_data, message="Target data preview")

# Read main sheet
if first_time_use:
  target_data.to_excel(
    f"{main_path}/{confirm_main_file}",
    index=False
  )

main_file_path = main_path + "/" + confirm_main_file
if debug: print(f"\nMain file is at {main_file_path}")

main_data = pd.read_excel(main_file_path)
if debug: preview_sheet_data(main_data, message="Main data preview")

# TODO: Remove duplicate entries in target sheet

# Clean the names in main and target data
clean_data(
  filename=confirm_main_file,
  data=main_data,
  name_column=name_column,
  debug=debug
)
clean_data(
  filename=confirm_target_file,
  data=target_data,
  name_column=name_column,
  debug=debug
)

# Read sheet columns and create new date column in main sheet
current_date = create_date_col_in_main(
  main_data=main_data,
  target_data=target_data,
  debug=debug
)

if debug: print(f"Current date as string: {str(current_date)}")

# Compare from both sheets
main_names = set(main_data[name_column])
target_names = set(target_data[name_column])

new_names_from_target = list(target_names - main_names)
new_names_count = len(new_names_from_target)
have_new_names = new_names_count != 0
if have_new_names:
  print(f"{new_names_count} newcomers detected\n{new_names_from_target}")
  main_data = add_new_names_to_sheet(
    main_sheet=main_data, 
    target_sheet=target_data,
    new_names=new_names_from_target,
    name_column=name_column,
    current_date=current_date,
    debug=debug
  )

missing_names_from_main = list(main_names - target_names)
missing_names_count = len(missing_names_from_main)
have_missing_names = missing_names_count != 0
if have_missing_names:
  print(f"\n{missing_names_count} absentees detected\n")
  # Tick names of people who attend (mark attendance with timestamp, absent as 'Absent')
  main_data = mark_absentees(
    sheet=main_data,
    current_date=current_date, 
    absentees=missing_names_from_main,
    name_column=name_column,
    debug=debug
  )

sheets_in_sync = (new_names_count == 0) & (missing_names_count == 0)
if sheets_in_sync: print(f"Sheets are in sync. No newcomers or absentees")

# Tick names of people who attend (mark attendance with timestamp, absent as 'Absent')
attendee_list = list(main_names & target_names)
attendees_count = len(attendee_list)
have_attendees = attendees_count != 0
if have_attendees:
  print(f"\n{attendees_count} attendees detected\n")
  main_data = mark_attendees(
    main_sheet=main_data,
    target_sheet=target_data,
    current_date=current_date, 
    attendees=attendee_list,
    name_column=name_column,
    debug=debug
  )

## TODO need a method for error handling (incorrect names, different capitalizations)

# TODO: Converting headers to normal

# Removing unused columns
if "Timestamp" in main_data.columns: main_data = main_data.drop('Timestamp', axis=1)

# Making sure headers are strings
main_data.columns.astype(str)

# BUG: Fix name sorting not working
# Sorting names from A to Z
print()
sorted_data = sort_names(
    sheet=main_data, 
    name_column=name_column,
    debug=debug
)

if debug: print(f"Sorted data in run.py: {sorted_data}")

# Save newly generated main sheet
if len(confirm_main_file.split("-")) > 1:
  sorted_data.to_excel(
    f"{main_path}/{current_date}-{confirm_main_file.split("-", maxsplit=4)[3]}",
    index=False
  )
else:
  sorted_data.to_excel(
    f"{main_path}/{current_date}-{confirm_main_file}",
    index=False
  )
  
# Move sheet from target into processed
if not debug:
  shutil.move(
    src=target_file_path, 
    dst=processed_path
  )

# Get rid of initial copy of file
initial_copy_path = f"{main_path}/{confirm_main_file}"
if first_time_use:
  if os.path.exists(initial_copy_path):
    os.remove(initial_copy_path)
    if debug: print(f"{initial_copy_path} has been deleted.")
  else:
      print(f"{initial_copy_path} does not exist.")