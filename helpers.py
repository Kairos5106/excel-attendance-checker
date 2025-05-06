from datetime import datetime
import os
import sys
import threading
import time

import pandas as pd

## Trivial loading animation
def loading_animation(stop_event, message="Loading"):
    spinner = ["|", "/", "-", "\\"]
    animation_index = 0
    while not stop_event.is_set():
        sys.stdout.write(f"\r{message} {spinner[animation_index % len(spinner)]}")
        sys.stdout.flush()
        time.sleep(0.1)
        animation_index += 1
    sys.stdout.write(f"\r{message}. Done!\n")
    sys.stdout.flush()

def sample_run_with_loader():
    stop_event = threading.Event()  # ✅ Define stop_event here
    loader_thread = threading.Thread(target=loading_animation, args=(stop_event, "Processing"))
    
    loader_thread.start()
    time.sleep(2) 
    stop_event.set() # Trigger loading animation to stop
    loader_thread.join()  # Wait for the loader thread to finish

## Read sheet data
def preview_sheet_data(data, message="Data preview"):
    print(message)
    print(data)
    print()

## Helper function for user to select files
def confirm_file_selection(files, type):
    if (len(files) == 1):
       print(f"{files[0]} is selected as main sheet.\n")
       return files[0]

    result = ""
    file_options = files

    while True:
      file_index = 0
      while True:
        user_input = input(f"From the list, select a file as {type} (e.g. 1, 2, 3, 4, ...) (control + c to exit) ")
        if user_input == "":
            print(f"Read the instructions properly")
            continue
        file_index = int(user_input) - 1
        if (file_index < 0) | (file_index > len(file_options) - 1):
           print(f"Invalid selection. Try again silly ")
           continue
        break
      result = file_options[file_index]
      user_input = input(f"Are you sure you want to select {result}? (y/n) ")
      if user_input.lower() in ["y", ""]:
         print(f"{result} was selected as {type} file\n")
         break
      elif user_input == "n":
         continue
      else:
         print(f"Invalid response. Try again dummy ")
    
    return result

## Helper function to safely list files
def safe_listdir(path, label, verbose = False):
    try:
        files = [f for f in os.listdir(path) if not f.startswith(".")]

        if verbose:
          print(f"{len(files)} {label} file(s) detected")
          if len(files) > 1:
            print(f"** CAUTION - There are multiple {label} files detected **")
            for animation_index, file in enumerate(files):
              print(f"{animation_index + 1}: {file}")

        return files
    except FileNotFoundError:
        print(f"Error: The folder '{path}' does not exist.")
        return []
    except PermissionError:
        print(f"Error: You do not have permission to access the folder '{path}'.")
        return []
    except Exception as e:
        print(f"Unexpected error while accessing '{path}': {e}")
        return []
    
## Standardize names in main and target
def clean_main_data(main_data, name_column, debug):
    stop_event = threading.Event()  # ✅ Define stop_event here
    loader_thread = threading.Thread(target=loading_animation, args=(stop_event, "Standardizing names in main sheet"))
    
    loader_thread.start()

    # Logic
    main_data[name_column] = main_data[name_column].str.upper()

    if debug:
        preview_sheet_data(main_data, message="Main sheet preview after cleaning")

    stop_event.set() # Trigger loading animation to stop
    loader_thread.join()  # Wait for the loader thread to finish

def clean_target_data(target_data, name_column, debug):
    stop_event = threading.Event()  # ✅ Define stop_event here
    loader_thread = threading.Thread(target=loading_animation, args=(stop_event, "Standardizing names in target sheet"))
    
    loader_thread.start()
    
    # Logic
    target_data[name_column] = target_data[name_column].str.upper()

    if debug:
        preview_sheet_data(target_data, message="Target sheet preview after cleaning")

    stop_event.set() # Trigger loading animation to stop
    loader_thread.join()  # Wait for the loader thread to finish

## Prompts user to confirm target date
def confirm_target_date(target_data):
    target_data['Timestamp'] = pd.to_datetime(target_data['Timestamp'])
    extracted_date = target_data["Timestamp"].iloc[0].date() ### Take the first target row's date as input (ask user for confirmation)

    while True:
      user_input = input(f"{extracted_date}: Is this the date you want to mark attendance for? (type y for yes OR type a new date in YYYY-MM-DD format [e.g. 2025-05-12]) ").strip()

      if user_input.lower() in ["y", ""]:
          return extracted_date
        
      # Try parsing the custom date input
      try:
          custom_date = datetime.strptime(user_input, "%Y-%m-%d").date()
          return custom_date
      except ValueError:
          print(f"⚠️ Invalid date format. Haih ... Please enter in YYYY-MM-DD format, or type 'y' to confirm this date: {extracted_date}")

## Create new date column in main sheet
def create_date_col_in_main(main_data, target_data, debug):
    main_data_columns = main_data.columns.to_list()
    print(f"Main sheet columns: {main_data_columns}")

    target_data_columns = target_data.columns.to_list()
    print(f"Target sheet columns: {target_data_columns}")

    print()

    # Compare main and target sheets
    ## Main cols: maintain but add one extra column for attendance on said date
    current_date = confirm_target_date(target_data=target_data)
    print(f"Current date selected: {current_date}")
    main_data[current_date] = None

    if debug:
        preview_sheet_data(main_data, message=f"Added {current_date} as a new column to main sheet")

    return current_date

## TODO: Add new names from target to main
def add_new_names_to_sheet(
    reference_sheet,
    target_sheet, 
    new_names,
    name_column,
    debug
):
    stop_event = threading.Event()  # ✅ Define stop_event here
    loader_thread = threading.Thread(target=loading_animation, args=(stop_event, "Adding new names from target into main sheet"))
    
    loader_thread.start()

    # Logic for adding new names
    print(f"New names: {new_names}")

    new_rows = target_sheet[target_sheet[name_column].isin(new_names)]
    print(f"New rows: {new_rows}")

    if debug:
        preview_sheet_data(target_sheet, message="Main sheet before addition")
    
    updated_sheet = pd.concat([target_sheet, new_rows], ignore_index=True)

    if debug:
        preview_sheet_data(updated_sheet, message="Main sheet after addition")

    # TODO: sync_metadata(ref=reference_sheet, target=target_sheet)

    stop_event.set() # Trigger loading animation to stop
    loader_thread.join()  # Wait for the loader thread to finish

    return updated_sheet

## Sort names in a sheet
def sort_names(sheet, name_column):
    stop_event = threading.Event()  # ✅ Define stop_event here
    loader_thread = threading.Thread(target=loading_animation, args=(stop_event, "Processing"))
    
    loader_thread.start()

    # Sorting logic
    sorted_sheet = sheet.sort_values(name_column)

    preview_sheet_data(sorted_sheet, message="Main sheet after sorting names")

    stop_event.set() # Trigger loading animation to stop
    loader_thread.join()  # Wait for the loader thread to finish

    return sorted_sheet

## Mark absentees in main
def mark_absentees(sheet, current_date, absentees, name_column, debug):
    stop_event = threading.Event()  # ✅ Define stop_event here
    loader_thread = threading.Thread(target=loading_animation, args=(stop_event, "Marking absentees"))
    
    loader_thread.start()
    
    # Mark absentees logic
    updated_sheet = sheet
    for absentee in absentees:
        updated_sheet.loc[updated_sheet[name_column] == absentee, current_date] = "Absent"
        print(f"Marked {absentee} as absent")

    if debug:
        preview_sheet_data(updated_sheet, message="Main sheet after marking absentees")

    stop_event.set() # Trigger loading animation to stop
    loader_thread.join()  # Wait for the loader thread to finish

    return updated_sheet

## Mark attendees in main
def mark_attendees(sheet, current_date, attendees, name_column):
    stop_event = threading.Event()  # ✅ Define stop_event here
    loader_thread = threading.Thread(target=loading_animation, args=(stop_event, "Marking attendees"))
    
    loader_thread.start()
    
    # Mark attendees logic
    updated_sheet = sheet
    for attendee in attendees:
        updated_sheet.loc[updated_sheet[name_column] == attendee, current_date] = updated_sheet['Timestamp']
        print(f"Marked {attendee} as present")

    preview_sheet_data(updated_sheet, message="Main sheet after marking attendees")

    stop_event.set() # Trigger loading animation to stop
    loader_thread.join()  # Wait for the loader thread to finish

    return updated_sheet

# TODO: Sync metadata to sheet from reference sheet
def sync_metadata(ref, target):
    stop_event = threading.Event()  # ✅ Define stop_event here
    loader_thread = threading.Thread(target=loading_animation, args=(stop_event, "Syncing student metadata"))
    
    loader_thread.start()
    
    # Add logic here

    stop_event.set() # Trigger loading animation to stop
    loader_thread.join()  # Wait for the loader thread to finish

    return updated_sheet
