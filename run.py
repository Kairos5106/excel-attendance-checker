import pandas as pd
import os
import shutil

from helpers import (
    add_new_names_to_sheet,
    clean_main_data,
    clean_target_data,
    create_date_col_in_main,
    mark_absentees,
    mark_attendees,
    preview_sheet_data,
    confirm_file_selection,
    safe_listdir,
    sort_names,
)

# Config
name_column = "Nama Penuh (Nitih IC)" # name columns must be the same in both main and target sheets
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

# Read the main sheet
print(f"\nYour current path is {current_path}\n")

# Get file lists with error handling
main_files = safe_listdir(main_path, "main", verbose=True)
confirm_main_file = confirm_file_selection(
  files=main_files, 
  type="main"
)

processed_files = safe_listdir(processed_path, "processed", verbose=False)

target_files = safe_listdir(target_path, "target", verbose=True)
confirm_target_file = confirm_file_selection(
  files=target_files, 
  type="target"
)

# TODO (after finish main stuff): When first time entry, create new file from selected target and end program
first_time_use = len(main_files) == 0

# Read main sheet
main_file_path = main_path + "/" + confirm_main_file
print(f"\nMain file is at {main_file_path}")

main_data = pd.read_excel(main_file_path)
preview_sheet_data(main_data, message="Main data preview")

# Read target sheet
target_file_path = target_path + "/" + confirm_target_file
print(f"\nTarget file is at {target_file_path}")

target_data = pd.read_excel(target_file_path)
preview_sheet_data(target_data, message="Target data preview")

# TODO: Remove duplicate entries in target sheet

# Clean the names in main and target data
clean_main_data(
  main_data=main_data,
  name_column=name_column,
  debug=debug
)
clean_target_data(
  target_data=target_data,
  name_column=name_column,
  debug=debug
)

# Read sheet columns and create new date column in main sheet
current_date = create_date_col_in_main(
  main_data=main_data,
  target_data=target_data,
  debug=debug
)

# Compare from both sheets
main_names = set(main_data[name_column])
target_names = set(target_data[name_column])

new_names_from_target = list(target_names - main_names)
new_names_count = len(new_names_from_target)
have_new_names = new_names_count != 0
if have_new_names:
  print(f"{new_names_count} newcomers detected\n{new_names_from_target}")
  main_data = add_new_names_to_sheet(
    reference_sheet=target_data,
    target_sheet=main_data, 
    new_names=new_names_from_target,
    name_column=name_column,
    debug=debug
  )
  main_data = sort_names(
    sheet=main_data, 
    name_column=name_column
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
if sheets_in_sync:
  print(f"Sheets are in sync. No newcomers or absentees")

# Tick names of people who attend (mark attendance with timestamp, absent as 'Absent')
attendee_list = list(main_names & target_names)
attendees_count = len(attendee_list)
have_attendees = attendees_count != 0
if have_attendees:
  print(f"\n{attendees_count} attendees detected\n")
  main_data = mark_attendees(
    sheet=main_data,
    current_date=current_date, 
    attendees=attendee_list,
    name_column=name_column
  )

## TODO need a method for error handling (incorrect names, different capitalizations)

# Removing unused columns
main_data = main_data.drop('Timestamp', axis=1)

# Save newly generated main sheet
main_data.to_excel(
  f"{main_path}/{current_date}-{confirm_main_file}.xlsx",
  index=False
)

# TODO: Move sheet from target into processed
# shutil.move(
#   src=target_file_path, 
#   dst=processed_path
# )