import os
import pandas as pd
import logging
from datetime import datetime, timedelta
import random

# Configure logging
log_directory = os.path.join(os.path.dirname(__file__), "logging")
data_directory = os.path.join(os.path.dirname(__file__), "data")
history_directory = os.path.join(os.path.dirname(__file__), "history")

# Create directories if they do not exist
for directory in [log_directory, data_directory, history_directory]:
    os.makedirs(directory, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(os.path.join(log_directory, "script_events.log")),
        logging.FileHandler(os.path.join(log_directory, "assignment_data_log.csv")),
    ],
)

# Constants
NUM_WEEKS = 52

# Set the paths to the input and output files
team_list_path = os.path.join(data_directory, "team_list.xlsx")
assignment_path = os.path.join(data_directory, "assignments.xlsx")
assignment_data_log_path = os.path.join(log_directory, "assignment_data_log.csv")

def generate_schedule_weeks():
    # Get the current date
    current_date = datetime.now()

    # Calculate the start date (Monday of the current week)
    start_date = current_date - timedelta(days=current_date.weekday())

    # Initialize lists for start and end dates
    start_dates = []
    end_dates = []

    # Generate start and end dates for each week
    for week in range(NUM_WEEKS):
        end_date = start_date + timedelta(days=4)  # Assuming assignments are from Monday to Friday
        start_dates.append(start_date.strftime("%m-%d-%Y"))
        end_dates.append(end_date.strftime("%m-%d-%Y"))

        # Move to the next week
        start_date += timedelta(days=7)

    # Create a DataFrame with start and end dates
    schedule_df = pd.DataFrame({"start_date": start_dates, "end_date": end_dates})

    # Write the schedule to assignments.xlsx
    schedule_df.to_excel(assignment_path, index=False)

    logging.info(f"Schedule dates written to {assignment_path}")
    print(f"Schedule dates written to {assignment_path}.")


def read_employee_data(file_path):
    try:
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        logging.error(f"Error reading employee data: {e}")
        print(f"Error reading employee data: {e}")
        return None

def load_and_detect_changes(current_employee_data, previous_employee_data):
    logging.info("Load and detect changes in employee data")
    print("Loading and detecting changes in employee data...")
    if previous_employee_data is not None:
        changes = detect_changes(previous_employee_data, current_employee_data)
        if changes:
            logging.info("Changes detected in team_list.xlsx:")
            print("Changes detected in team_list.xlsx:")
            for change in changes:
                logging.info(change)
                print(change)

def detect_changes(old_data, new_data):
    changes = []

    for index, row in old_data.iterrows():
        old_row = row.to_dict()
        new_row = new_data[new_data["Name"] == old_row["Name"]].to_dict("records")[0]

        for key, old_value in old_row.items():
            new_value = new_row[key]
            if old_value != new_value:
                changes.append(
                    f"Change in {key} for employee {old_row['Name']}: {old_value} -> {new_value}"
                )

    return changes

def log_activity(activity_description):
    formatted_timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    logging.info(f"{formatted_timestamp} - {activity_description}")
    print(f"{formatted_timestamp} - {activity_description}")

def backup_existing_assignments():
    logging.info("Back up existing assignments")
    print("Backing up existing assignments...")
    if os.path.exists(assignment_data_log_path):
        timestamp_str = datetime.now().strftime("%m-%d-%Y_%H-%M-%S")
        new_file_name = os.path.join(history_directory, f"assignment_data_log_as_of_{timestamp_str}.csv")

        os.rename(assignment_data_log_path, new_file_name)
        logging.info(f"Back up {assignment_path} to {new_file_name}")
        print(f"Assignment log backed up to {new_file_name}")

        # Keep only the last 3 backups, delete older ones
        delete_old_backups(history_directory, "assignment_data_log", keep_latest=3)
    else:
        logging.info("No existing assignment data log to back up.")
        print("No existing assignment data log to back up.")

def delete_old_backups(directory, base_name, keep_latest):
    logging.info(f"Delete old backups in {directory}")
    print(f"Deleting old backups in {directory}...")
    files = [f for f in os.listdir(directory) if f.startswith(f"{base_name}_as_of_")]
    files.sort(key=lambda x: os.path.getctime(os.path.join(directory, x)), reverse=True)

    # Delete older backups, keep only the last 'keep_latest'
    for old_file in files[keep_latest:]:
        os.remove(os.path.join(directory, old_file))
        logging.info(f"Retain maximum of {keep_latest} saved logs. Old {base_name} backup deleted: {old_file}")
        print(f"Retain maximum of {keep_latest} saved logs. Old {base_name} backup deleted.")

def main():
    logging.info("Start script execution")
    print("Starting script execution...")
    current_employee_data = read_employee_data(team_list_path)
    previous_employee_data = read_employee_data(team_list_path)

    if current_employee_data is not None:
        load_and_detect_changes(current_employee_data, previous_employee_data)

        backup_existing_assignments()

        generate_schedule_weeks()

        logging.info("Script execution completed.\n")
        print("Script execution completed.\n")
    else:
        logging.error("Exiting program due to errors.\n")
        print("Exiting program due to errors.\n")

if __name__ == "__main__":
    main()
