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
Some issues exist with inconsistent time formats. For get results edit input file in excel

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
                
                # Default start and end times for error cases
                start_time = "07:50:00 AM"
                end_time = "07:55:00 AM"
                description = ""
                key_name = str(row['Experiment Name'])

                try:
                    # Attempt to find period and set times
                    period_index = int(row['Period']) - 1
                    
                    # A) Set end time to 5 minutes after start time
                    start_time_str = times_table[period_index][1]
                    start_datetime_obj = datetime.datetime.strptime(start_time_str, "%I:%M:%S %p")
                    end_datetime_obj = start_datetime_obj + datetime.timedelta(minutes=5)
                    end_time_str = end_datetime_obj.strftime("%I:%M:%S %p")
                    
                    start_time = start_time_str
                    end_time = end_time_str

                except (ValueError, IndexError):
                    # B) If an error occurs, use the default 7:50 AM time
                    description = "Error: Period number not found or invalid."
                    print(f"Skipping row due to invalid period data for '{key_name}'. Placing appointment at 7:50 AM. Full row: {row}")
                
                # If information is needed in the body text of the event, add it here.
                # If a description was already set due to an error, don't overwrite it.
                if not description:
                    description = ""

                intermediate_array.append(
                    [key_name, row['Date'], start_time, row['Date'], end_time,
                     '', '', '', '', '',
                     '', '', '', '', 'RiskAssess'
                        , '', row['Room'], description, '',
                     ]
                )
            except (KeyError, ValueError, IndexError) as e:
                # Catch issues with other fields, like 'Experiment Name' or 'Date'
                print(f"Skipping row due to missing or invalid data: {row}. Error: {e}")
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
