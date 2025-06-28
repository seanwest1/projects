### Riskassess CSV to Outlook Calendar Converter

### Project Description
* This Python script automates the conversion of CSV files downloaded from `www.riskassess.com.au` into a format that can be easily imported into Microsoft Outlook Calendar. It streamlines the process of adding your lab or experiment schedules to your personal calendar making them much easier to visualise.

### Features
* Converts Riskassess.com.au CSV data (Experiment Name, Date, Period, Room) into Outlook-compatible calendar fields (Subject, Start Date, Start Time, End Date, End Time, Location, Categories).
* Adjusts end times for specific schedule exceptions (e.g., Period 5 on a Wednesday).
* Provides a user-friendly interface for selecting the input CSV file from the current directory.
* Generates an `output.csv` file ready for direct import into Outlook Calendar.
* Includes basic error handling for malformed rows in the input CSV to prevent crashes and provide informative messages.

### How to Use
1.  **Download the Script:** Save the Python script (e.g., `riskassess_converter.py`) to your local machine.
2.  **Obtain Riskassess CSV:** Log in to `www.riskassess.com.au` and download your desired schedule as a CSV file.
3.  **Place the CSV:** Move the downloaded CSV file into the **same directory** as the Python script.
4.  **Rename (Optional but Recommended):** For convenience, you can rename your downloaded CSV file to `lab_schedule.csv`. If not, you'll need to select it during script execution.
5.  **Run the Script:**
(If python is configured correctly you may be able to just run the script on windows as though it were a file or program)
    * Open your terminal or command prompt.
    * Navigate to the directory where you saved the script.
    * Execute the script using: `python riskassess_converter.py`
6.  **Follow Prompts:**
    * The script will ask you to "Choose a file (press enter to select from list) or use 'lab_schedule.csv' (press space then enter):".
        * Press `Enter` to see a numbered list of CSV files in the directory and select one by typing its corresponding number.
        * Press `Space` then `Enter` to automatically use a file named `lab_schedule.csv` (if it exists).
7.  **Output File:** A new file named `output.csv` will be created in the same directory.

### Importing into Outlook Calendar
1.  Open Outlook.
2.  Go to the Calendar view.
3.  Navigate to `File` > `Open & Export` > `Import/Export`.
4.  Choose `Import from another program or file` and click `Next`.
5.  Select `Comma Separated Values` and click `Next`.
6.  Browse to your `output.csv` file and click `Next`.
7.  Select the destination folder (usually your Calendar) and click `Next`.
8.  Click `Map Custom Fields...`. Drag the fields from the left column (from `output.csv`) to the corresponding Outlook fields on the right. Key mappings include:
    * `Subject` -> `Subject`
    * `Start Date` -> `Start Date`
    * `Start Time` -> `Start Time`
    * `End Date` -> `End Date`
    * `End Time` -> `End Time`
    * `Location` -> `Location`
    * `Description` -> `Description`
    * `Categories` -> `Categories`
9.  Click `OK`, then `Finish`. Your events should now appear in your Outlook Calendar.

### Input CSV Requirements/Assumptions
* The input CSV must be obtained from `www.riskassess.com.au`.
* Expected column headers include (but are not limited to): `'Experiment Name'`, `'Date'`, `'Period'`, and `'Room'`.
* **Date Format:** The script assumes the 'Date' column in the input CSV is in `YYYY-MM-DD` format. If your Riskassess export uses a different format (e.g., `DD/MM/YYYY`), you may need to modify the `datetime.datetime.strptime` line in the script or pre-process your CSV.

* **Time Consistency:** The script attempts to handle standard periods and a specific Wednesday Period 5 exception. Inconsistent time formats in the original Riskassess CSV might lead to errors. It is recommended to check and, if necessary, edit the input CSV in a text editor (not Excel) to ensure consistent time formats before running the script. Excel changes the date format.

### Error Handling
* The script includes basic error handling for individual rows in the input CSV. If a row cannot be processed due to missing or malformed data (e.g., a missing 'Period' or invalid date), it will print an error message to the console for that specific row and skip it, allowing the rest of the file to be processed.

### Contributing
* Feel free to fork this repository, open issues, or submit pull requests if you have suggestions for improvements or encounter bugs.
