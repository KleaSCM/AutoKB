Overview
AutoKB is a Python-based automation tool designed to streamline repetitive tasks such as pasting pre-defined text content and processing CSV files into Excel format. The tool leverages keybindings to trigger these actions, making it a highly efficient solution for automating workflows.

Features
Text Pasting Automation: Assign key combinations to paste content from text files directly into active applications.
CSV to Excel Conversion: Automatically process CSV files, delete specified columns, and save the results as styled Excel files.
Custom Keybindings: Easily configurable keybindings for various automation tasks.
Real-time Hotkey Monitoring: Continuously monitor and respond to hotkey events until the Esc key is pressed.
Installation
Clone the Repository:

bash
Copy code
git clone https://github.com/your-username/AutoKB.git
cd AutoKB
Install Dependencies:
Ensure you have Python 3.x installed. Install the required Python libraries using pip:

bash
Copy code
pip install keyboard pyperclip pandas pyautogui openpyxl
Configuration
Edit the keybindings dictionary in the script to customize the key combinations and the associated file paths. Example:

python
Copy code
keybindings = {
    'Ctrl+shift+1': r'C:\path\to\your\file1.txt',
    'Ctrl+shift+2': r'C:\path\to\your\file2.txt',
    # Add more keybindings as needed
}
Usage
Run the Script:

bash
Copy code
python AutoKB.py
Activate Hotkeys:

Use the defined key combinations to trigger text pasting or CSV processing.
Press Esc to stop the script.
Keybindings
The following key combinations are predefined for this project:

Text Files:

Ctrl+shift+1: Paste content from DNAI Cultural Business.txt
Ctrl+shift+2: Paste content from DNAI Sorry Business.txt
Ctrl+shift+3: Paste content from DNAV Community Unrest.txt
Ctrl+shift+4: Paste content from DNAV Medical.txt
Ctrl+shift+5: Paste content from DNAV Non-compelable.txt
Ctrl+shift+6: Paste content from DNAV Work-related.txt
Ctrl+shift+7: Paste content from DNAV Voluntary.txt
Ctrl+shift+8: Paste content from DNAV No Staff in Community.txt
Ctrl+shift+9: Paste content from DNAV Flooding.txt
Ctrl+shift+0: Paste content from DNAI Contactable.txt
Ctrl+shift+-: Paste content from DNAI Non-contactable.txt
Ctrl+shift+=: Paste content from DNAI Income Phone.txt
Ctrl+shift+ : Paste content from DNAI Income.txt`
CSV Processing:

Ctrl+shift+alt+0: Process and convert 216-CDPJobPlanMonitoring-20240428.csv to Excel.
CSV Processing Details
The script processes CSV files by deleting specific columns and saving the results in an Excel file with the following enhancements:

Thin borders around all cells.
Thick borders, blue background, and bold text for the first row.
Development
Feel free to contribute to this project by:

Adding new features.
Improving the existing code.
Reporting issues and suggesting improvements.
License
This project is licensed under the MIT License.

Acknowledgements
Special thanks to the open-source community for the following libraries:

keyboard
pyperclip
pandas
pyautogui
openpyxl
