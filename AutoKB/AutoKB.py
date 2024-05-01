# import pyautogui
# import keyboard
# import pyperclip 

# # Define a dictionary to store key combinations
# keybindings = {
#     # Resulting comments scripts
#     'ctrl+shift+1': r'C:\Users\Kliea\Documents\Dev\Proj\Python\Scripts\DNAI Sorry Buisiness.txt',
#     'ctrl+shift+2': r'C:\Users\Kliea\Documents\Dev\Proj\Python\Scripts\DNAI Cultural Buisiness.txt',
#     'ctrl+shift+3': r'C:\Users\Kliea\Documents\Dev\Proj\Python\Scripts\DNAV Non-compelable.txt',
#     'ctrl+shift+4': r'C:\Users\Kliea\Documents\Dev\Proj\Python\Scripts\DNAV Work related.txt',
#     'ctrl+shift+5': r'C:\Users\Kliea\Documents\Dev\Proj\Python\Scripts\DNAV community Unrest.txt',
#     'ctrl+shift+6': r'C:\Users\Kliea\Documents\Dev\Proj\Python\Scripts\DNAV Medical.txt',
#     # Add other keybindings
# }

# def paste_text(filepath):
#     # Open the text document and copy the content to the clipboard
#     with open(filepath, 'r') as file:
#         content = file.read().strip()  # Strip trailing newline characters
#     pyperclip.copy(content)  # Copy content to clipboard
#     pyautogui.hotkey('ctrl', 'v')  # Paste content from clipboard

# def on_hotkey(hotkey):
#     # Get the file path associated with the hotkey
#     filepath = keybindings.get(hotkey)
#     if filepath:
#         # Paste content from the txt doc into the active window
#         paste_text(filepath)

# # Register the hotkey callback function for each hotkey combination
# for hotkey in keybindings.keys():
#     keyboard.add_hotkey(hotkey, on_hotkey, args=(hotkey,))

# # Start the keyboard listener for hotkey events
# keyboard.wait('esc')  # Press 'Esc' to stop


#######################################################################################



# import pyautogui
# import keyboard
# import pyperclip
# import pandas as pd

# # Define a dictionary to store key combinations
# keybindings = {
#     # Resulting comments scripts
#     'ctrl+shift+1': r'C:\Users\Kliea\Documents\Dev\Proj\Python\Scripts\DNAI Sorry Buisiness.txt',
#     'ctrl+shift+2': r'C:\Users\Kliea\Documents\Dev\Proj\Python\Scripts\DNAI Cultural Buisiness.txt',
#     'ctrl+shift+3': r'C:\Users\Kliea\Documents\Dev\Proj\Python\Scripts\DNAV Non-compelable.txt',
#     'ctrl+shift+4': r'C:\Users\Kliea\Documents\Dev\Proj\Python\Scripts\DNAV Work related.txt',
#     'ctrl+shift+5': r'C:\Users\Kliea\Documents\Dev\Proj\Python\Scripts\DNAV community Unrest.txt',
#     'ctrl+shift+6': r'C:\Users\Kliea\Documents\Dev\Proj\Python\Scripts\DNAV Medical.txt',
#     # Add other keybindings
#     'ctrl+shift+alt+0': r'C:\Users\Kliea\Documents\Dev\Proj\Python\Reports\216-CDPJobPlanMonitoring-20240428.csv'
# }

# def paste_text(filepath):
#     # Open the text document and copy the content to the clipboard
#     with open(filepath, 'r') as file:
#         content = file.read().strip()  # Strip trailing newline characters
#     pyperclip.copy(content)  # Copy content to clipboard
#     pyautogui.hotkey('ctrl', 'v')  # Paste content from clipboard

# def process_csv_and_save_as_xlsx(file_path, columns_to_delete, output_file):
#     df = pd.read_csv(file_path)

#     # Delete specified columns
#     df.drop(columns_to_delete, axis=1, inplace=True)

#     # Save as XLSX
#     df.to_excel(output_file, index=False)

# def on_hotkey(hotkey):
#     # Get the file path associated with the hotkey
#     filepath = keybindings.get(hotkey)
#     if filepath:
#         if filepath.endswith('.txt'):
#             # Paste content from the txt doc into the active window
#             paste_text(filepath)
#         elif filepath.endswith('.csv'):
#             # Process the CSV file and save as XLSX
#             # Example: Delete specified columns
#             columns_to_delete = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'I', 'M', 'Q', 'R', 'S', 'T', 'U', 'Y', 'Z',
#                                  'AB', 'AD', 'AE', 'AH', 'AI', 'AJ', 'AL', 'AM', 'AO', 'AN', 'AP', 'AQ',
#                                  'AS', 'AT', 'AV', 'AW', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG',
#                                  'BH', 'BI', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU',
#                                  'BV', 'BW', 'BX', 'BY', 'BZ', 'CA', 'CC', 'CD', 'CE', 'CF', 'CG', 'CJ', 'CK', 'CM',
#                                  'CN', 'CQ', 'CR', 'CU', 'CV', 'CW', 'CY', 'CZ', 'DA', 'DC', 'DD', 'DE', 'DF', 'DI']
#             output_file = 'output.xlsx'  # Output XLSX file
#             process_csv_and_save_as_xlsx(filepath, columns_to_delete, output_file)

# # Register the hotkey callback function for each hotkey combination
# for hotkey in keybindings.keys():
#     keyboard.add_hotkey(hotkey, on_hotkey, args=(hotkey,))

# # Start the keyboard listener for hotkey events
# keyboard.wait('esc')  # Press 'Esc' to stop


import pyautogui
import keyboard
import pyperclip
import pandas as pd

# Define a dictionary to store key combinations
keybindings = {
    # Resulting comments scripts
    'ctrl+shift+1': r'C:\Users\Kliea\Documents\Dev\Proj\Python\Scripts\DNAI Sorry Buisiness.txt',
    'ctrl+shift+2': r'C:\Users\Kliea\Documents\Dev\Proj\Python\Scripts\DNAI Cultural Buisiness.txt',
    'ctrl+shift+3': r'C:\Users\Kliea\Documents\Dev\Proj\Python\Scripts\DNAV Non-compelable.txt',
    'ctrl+shift+4': r'C:\Users\Kliea\Documents\Dev\Proj\Python\Scripts\DNAV Work related.txt',
    'ctrl+shift+5': r'C:\Users\Kliea\Documents\Dev\Proj\Python\Scripts\DNAV community Unrest.txt',
    'ctrl+shift+6': r'C:\Users\Kliea\Documents\Dev\Proj\Python\Scripts\DNAV Medical.txt',
    # Add other keybindings
    'ctrl+shift+alt+0': r'C:\Users\Kliea\Documents\Dev\Proj\Python\Reports\216-CDPJobPlanMonitoring-20240428.csv'
}

def paste_text(filepath):
    # Open the text document and copy the content to the clipboard
    with open(filepath, 'r') as file:
        content = file.read().strip()  # Strip trailing newline characters
    pyperclip.copy(content)  # Copy content to clipboard
    pyautogui.hotkey('ctrl', 'v')  # Paste content from clipboard

def process_csv_and_save_as_xlsx(file_path, columns_to_delete, output_file):
    df = pd.read_csv(file_path)

    # Delete specified columns
    df.drop(columns=df.columns[columns_to_delete], inplace=True)

    # Save as XLSX
    df.to_excel(output_file, index=False)

def on_hotkey(hotkey):
    # Get the file path associated with the hotkey
    filepath = keybindings.get(hotkey)
    if filepath:
        if filepath.endswith('.txt'):
            # Paste content from the txt doc into the active window
            paste_text(filepath)
        elif filepath.endswith('.csv'):
            # Process the CSV file and save as XLSX
            columns_to_delete = [ 0, 1, 2, 3, 4, 5, 6, 8, 12, 15, 16, 17, 81, 19,
                                20, 25, 27, 29, 30, 33, 34, 42, 44, 45, 47, 48, 50,
                                51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 62, 63,
                                64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75,
                                76, 77, 78, 80, 81, 82, 83, 84, 86, 87, 88, 89,
                                90, 91, 94, 95, 99, 100, 102, 103, 104, 106, 107, 108, 109  ]
            output_file = 'output.xlsx'  # Output XLSX file
            process_csv_and_save_as_xlsx(filepath, columns_to_delete, output_file)

# Register the hotkey callback function for each hotkey combination
for hotkey in keybindings.keys():
    keyboard.add_hotkey(hotkey, on_hotkey, args=(hotkey,))

# Start the keyboard listener for hotkey events
keyboard.wait('esc')  # Press 'Esc' to stop
