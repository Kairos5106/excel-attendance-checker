import pandas as pd
import os

# Configuration
verbose = False

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
print(f"Your current path is {current_path}")

# Helper function to safely list files
def safe_listdir(path, label):
    try:
        files = [f for f in os.listdir(path) if not f.startswith(".")]
        print(f"{len(files)} {label} file(s) detected")

        if len(files) > 1:
          print(f"** CAUTION - There are multiple {label} files detected **")
          for index, file in enumerate(files):
            print(f"{index + 1}: {file}")

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

# Get file lists with error handling
main_files = safe_listdir(main_path, "main")
processed_files = safe_listdir(processed_path, "processed")
target_files = safe_listdir(target_path, "target")

## Helper function for user to select files
def confirm_file_selection(type):
    if (len(main_files) == 1) & (type == "main"):
       return main_files[0]
    if (len(target_files) == 1) & (type == "target"):
       return target_files[0]

    result = ""
    file_options = []

    if type == "main":
        file_options = main_files
    if type == "target":
        file_options = target_files

    while True:
      file_number = input(f"From the list, select a file as {type} (e.g. 1, 2, 3, 4, ...) (control + c to exit) ")
      file_index = int(file_number) - 1
      result = file_options[file_index]
      confirm_file_selection = input(f"Are you sure you want to select {result}? (y/n) ")
      if confirm_file_selection == "n":
         continue
      else:
         break
    
    return result

confirm_main_file = confirm_file_selection(type="main")
confirm_target_file = confirm_file_selection(type="target")

# Read main sheet
main_file_path = main_path + "/" + confirm_main_file
print(f"\nMain file is at {main_file_path}")

main_data = pd.read_excel(main_file_path)
print("Main data preview:")
print(main_data)
print()

# Read target sheet
target_file_path = target_path + "/" + confirm_target_file
print(f"\nTarget file is at {target_file_path}")

target_data = pd.read_excel(target_file_path)
print("Target data preview:")
print(target_data)
print()

# Read sheet columns
main_data_columns = main_data.columns.to_list()
print(f"Main sheet columns: {main_data_columns}")

target_data_columns = target_data.columns.to_list()
print(f"Target sheet columns: {target_data_columns}")

# Compare main and target columns
## Main cols: maintain but add one extra column for attendance on said date
### Take today's date as input (ask user for confirmation)
### If target date not in main sheet, create new column in main

### map names in target to main
## need a method for error handling (incorrect names, different capitalizations)

# move sheet from target into processed
# find some way to selectively process one sheet out of multiple sheets