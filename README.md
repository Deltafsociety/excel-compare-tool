
# Excel File Comparator with Highlighting

This Python application provides a user-friendly Graphical User Interface (GUI) to compare two Excel files (.xlsx or .xls) and highlight matching cell values. It offers two types of highlighting:

    Yellow Highlight: For all matching cell values (text or numeric).

    Red Highlight: Specifically for matching cell values that are identified as numbers.
The application processes all sheets within the selected Excel files and generates new highlighted versions, leaving your original files untouched.



Features

    GUI Interface: Easy file selection and operation using tkinter.

    Comprehensive Comparison: Compares data across all sheets within both Excel files.

    Dual Highlighting: Distinguishes between general matches (yellow) and shared numeric matches (red).

    Non-Destructive: Creates new highlighted Excel files, preserving your original data.

    Status Updates: Provides real-time feedback on the comparison process.

Installation

To run this application, you need Python installed on your system.
You also need the pandas and openpyxl libraries.

    Clone or Download the Repository:
    (Assuming you have the Python script excel_comparator_gui.py and requirements.txt.)

    Install Dependencies:
    Use the package manager pip to install the required Python libraries. Navigate to the directory where you saved the files in your terminal or command prompt and run the following command:

    pip install -r requirements.txt

    The requirements.txt file contains:

    pandas
    openpyxl

Usage

    Run the Application:
    Execute the Python script from your terminal:

    python excel_comparator_gui.py

    Select Excel Files:

        A GUI window titled "ROYAL CLASSIFICATION SOCIETY" will appear.

        Click the "Browse" button next to "Excel File 1" to select your first Excel file.

        Click the "Browse" button next to "Excel File 2" to select your second Excel file.

    Start Comparison:

        Click the "Compare and Highlight" button.

        The "Status" area will display the progress of the comparison.

        Once completed, a success message box will pop up.

    View Results:

        New highlighted Excel files will be generated in a subdirectory named highlighted_excel_files (or a custom directory if specified in the code).

        The new files will have _highlighted appended to their original names (e.g., your_file1_highlighted.xlsx).

Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.


## License

[GPL3](https://www.gnu.org/licenses/gpl-3.0.en.html)


## Tech Stack

Python3

