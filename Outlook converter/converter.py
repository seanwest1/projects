```python
""" Convert Riskassess.com.au CSV to CSV for outlook

This script takes a CSV file downloaded from www.riskassess.com.au
and converts it into a format that can be imported into outlook
calendar


Usage:
Place the downloaded CSV file to the same folder as the script,
rename it to "lab_schedule" or select it from the dialog in the script.
Run the script and it will rearrange the data into dates and times and
names that are compatible with Outlook.

Input:
CSV file from riskassess.com.
Some issues exist with inconsistant time formats. For get results edit input file in excel

Output:
CSV, readily importable to outlook calendar.

"""

import csv
import datetime
import os


def load_and_convert_CSV(riskassess_CSV_file_name):
    """ Imports the data from a CSV file and arranges it in an array """
    intermediate_array = []  # the array to parse/pass between functions
    with open(riskassess_CSV_file_name, newline='') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            try:
                # Period times
                times_table = [
                    [1, "08:45:00 AM", "09:30:00 AM"],
                    [2, "09:30:00 AM", "10:50:00 AM"],
                    [3, "11:10:00 AM", "12:35:00 PM"],
                    [4, "12:35:00 PM", "01:15:00 PM"],
                    [5, "01:55:00 PM", "03:15:00 PM"],
                    [6, "01:55:00 PM", "02:35:00 PM"]
                ]
                # to make the following less messy, convert periods to indexes here
                period_index = int(row['Period']) - 1
                time_look_up_shortened = times_table[period_index]

                """ Take a date in dd/mm/yyyy string and make it useable by the
                calendar module to see if it's a wednesday


                """
                # see which day of the week it is. Monday being 0 and Wednesday being 2
                # Assuming the date format in the CSV is YYYY-MM-DD based on the original strptime format
                day_number = datetime.datetime.strptime(row['Date'], "%Y-%m-%d").weekday()

                period_number_checker = int(row['Period'])

                if (day_number == 2) and (period_number_checker == 5):
                    # if a class is at period 5 on a Wednesday, use 2:35pm as end time
                    start_time = time_look_up_shortened[1]
                    end_time = times_table[5][2]  # This is times_table[index 5, end time] which is the same as Period 6 end time
                else:
                    # otherwise look it up in the table
                    start_time = time_look_up_shortened[1]
                    end_time = time_look_up_shortened[2]

                # If information is needed in the body text of the event, put it here.
                description = ""

                key_name = str(row['Experiment Name'])
                intermediate_array.append(
                    [key_name, row['Date'], start_time, row['Date'], end_time,
                     '', '', '', '', '',
                     '', '', '', '', 'RiskAssess'
                        , '', row['Room'], description, '',
                     ]
                )
            except (KeyError, ValueError, IndexError) as e:
                print(f"Skipping row due to data issue: {row}. Error: {e}")
                print(f"Please check 'Experiment Name': '{row.get('Experiment Name', 'N/A')}', 'Period': '{row.get('Period', 'N/A')}', and 'Date': '{row.get('Date', 'N/A')}' for errors.")
                continue # Skip to the next row if an error occurs

    return intermediate_array


def outlook_CSV_writer(converted_timetable_array):
    """Take the array created earlier and  write it to a new CSV output file """

    with open('output.csv', 'w', newline="") as f:

        ics_format_fields = ["Subject", "Start Date", "Start Time",
                             "End Date", "End Time", "All day event",
                             "Reminder on/off", "Reminder Date", "Reminder Time",
                             "Meeting Organizer", "Required Attendees", "Optional Attendees",
                             "Meeting Resources", "Billing Information", "Categories", "Description",
                             "Location", "Mileage", "Priority", "Private", "Sensitivity", "Show time as"]

        write = csv.writer(f)
        write.writerow(ics_format_fields)
        write.writerows(converted_timetable_array)


def path_view_select():
    """ Allow the user to select a CSV to use in the script """

    while True: # Use a while loop instead of recursion
        try:
            array_of_files_in_path = os.listdir(path='.')
            file_index = 0

            print("\nAvailable files in current directory:")
            for file in array_of_files_in_path:
                file_index += 1
                print(f"{file_index} : {file}")

            selection_str = input(f"Please enter the corresponding number to your file: ")
            selection = int(selection_str) # Attempt to convert input to integer

            if 1 <= selection <= len(array_of_files_in_path):
                chosen_file = array_of_files_in_path[selection - 1]
                print(f"You chose: {chosen_file}")

                if ".csv" in chosen_file.lower(): # Check for .csv extension case-insensitively
                    return chosen_file
                else:
                    print("Selected file is not a CSV. Please select a CSV file.")
            else:
                print("Invalid number. Please enter a number from the list.")

        except ValueError:
            print("Invalid input. Please enter a number.")
        except IndexError:
            print("Invalid selection. The number is out of range.")
        except Exception as e: # Catch any other unexpected errors
            print(f"An unexpected error occurred: {e}")
        print("Press Ctrl + C to exit, or select another file.")


user_choice = input("Choose a file (press enter to select from list) or use 'lab_schedule.csv' (press space then enter): ")
if user_choice == "":
    # User pressed Enter, so let them select a file
    selected_file = path_view_select()
    if selected_file: # Ensure a file was actually selected before proceeding
        outlook_CSV_writer(load_and_convert_CSV(selected_file))
else:
    # User pressed space then enter, or typed something else, defaulting to lab_schedule.csv
    outlook_CSV_writer(load_and_convert_CSV('lab_schedule.csv'))
```
