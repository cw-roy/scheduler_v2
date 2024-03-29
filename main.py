import os
import shutil
import pandas as pd
import logging
from datetime import datetime, timedelta
import random
from collections import deque

# Set up constants, folder structure and logging
NUM_WEEKS = 52
log_directory = os.path.join(os.path.dirname(__file__), "logging")
data_directory = os.path.join(os.path.dirname(__file__), "data")
history_directory = os.path.join(os.path.dirname(__file__), "history")

for directory in [log_directory, data_directory, history_directory]:
    os.makedirs(directory, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(os.path.join(log_directory, "script_events.log")),
    ],
)

# Set the paths to the input and output files
team_list_path = os.path.join(data_directory, "team_list.xlsx")
assignment_path = os.path.join(data_directory, "assignments.xlsx")


def generate_schedule_weeks():
    current_date = datetime.now()
    start_date = current_date - timedelta(days=current_date.weekday())

    # Check if the start_date is not a Monday (weekday() returns 0 for Monday)
    if start_date.weekday() != 0:
        start_date = start_date - timedelta(days=start_date.weekday())

    start_dates, end_dates = [], []
    for week in range(NUM_WEEKS):
        end_date = start_date + timedelta(
            days=4
        )  # Assuming assignments are from Monday to Friday
        start_dates.append(start_date.strftime("%m-%d-%Y"))
        end_dates.append(end_date.strftime("%m-%d-%Y"))
        start_date += timedelta(days=7)

    schedule_df = pd.DataFrame({"start_date": start_dates, "end_date": end_dates})

    with pd.ExcelWriter(assignment_path, engine="xlsxwriter") as writer:
        schedule_df.to_excel(writer, sheet_name="Weeks", index=False)

    logging.info(f"Schedule dates written to {assignment_path}")
    # print(f"Schedule dates written to {assignment_path}.")


def validate_team_list(team_df):
    # Validate headers
    expected_headers = ["Name", "Email", "Available"]
    if list(team_df.columns) != expected_headers:
        logging.error(
            f"Invalid headers in team_list.xlsx. Expected: {expected_headers}"
        )
        return False

    # Validate email format
    if not team_df["Email"].astype(str).apply(lambda x: "@" in x and "." in x).all():
        logging.error("Invalid email format in the 'Email' column of team_list.xlsx")
        return False

    # Validate 'Available' column values
    valid_available_values = ["yes", "no"]
    if (
        not team_df["Available"]
        .astype(str)
        .str.lower()
        .isin(valid_available_values)
        .all()
    ):
        logging.error(
            "Invalid values in the 'Available' column of team_list.xlsx. Allowed values: 'yes' or 'no'"
        )
        return False

    return True


def read_employee_data(file_path):
    try:
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        logging.error(f"Error reading employee data: {e}")
        return None


def load_and_detect_changes(current_employee_data, previous_employee_data):
    logging.info("Load and detect changes in employee data")
    if previous_employee_data is not None:
        changes = detect_changes(previous_employee_data, current_employee_data)
        if changes:
            logging.info("Changes detected in team_list.xlsx:")
            for change in changes:
                logging.info(change)


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


def backup_existing_assignments():
    logging.info("Back up existing assignments")
    if os.path.exists(assignment_path):
        timestamp_str = datetime.now().strftime("%m-%d-%Y_%H-%M-%S")
        new_file_name = os.path.join(
            history_directory, f"assignments_backup_as_of_{timestamp_str}.xlsx"
        )

        shutil.copy2(assignment_path, new_file_name)
        logging.info(f"Back up {assignment_path} to {new_file_name}")

        # Keep only the last 3 backups, delete older ones
        delete_old_backups(history_directory, "assignments_backup", keep_latest=3)
    else:
        logging.info("No existing assignment data log to back up.")


def backup_existing_team_list():
    logging.info("Back up existing team list")
    if os.path.exists(team_list_path):
        timestamp_str = datetime.now().strftime("%m-%d-%Y_%H-%M-%S")
        new_file_name = os.path.join(
            history_directory, f"team_list_backup_as_of_{timestamp_str}.xlsx"
        )

        shutil.copy2(team_list_path, new_file_name)
        logging.info(f"Back up {team_list_path} to {new_file_name}")

        # Keep only the last 3 backups, delete older ones
        delete_old_backups(history_directory, "team_list_backup", keep_latest=3)
    else:
        logging.info("No existing team list to back up.")


def delete_old_backups(directory, base_name, keep_latest):
    logging.info(f"Delete old backups in {directory}")
    files = [f for f in os.listdir(directory) if f.startswith(f"{base_name}_as_of_")]
    files.sort(key=lambda x: os.path.getctime(os.path.join(directory, x)), reverse=True)

    # Delete older backups, keep only the last 'keep_latest'
    for old_file in files[keep_latest:]:
        os.remove(os.path.join(directory, old_file))
        logging.info(
            f"Retain maximum of {keep_latest} saved logs. Old {base_name} backup deleted: {old_file}"
        )


def generate_random_pairs(agents):
    random.shuffle(agents)
    pairs = []
    while agents:
        agent1 = agents.pop()
        agent2 = agents.pop() if agents else None
        pairs.append((agent1, agent2))
    return pairs


def count_assignments(schedule_df, team_df, counts_df=None):
    if counts_df is None:
        counts_df = pd.DataFrame(columns=["Name", "Assignments"])

    for index, row in schedule_df.iterrows():
        agent1 = row["Agent1"]
        agent2 = row["Agent2"]

        if pd.notna(agent1):
            counts_df = update_counts(agent1, counts_df)
        if pd.notna(agent2):
            counts_df = update_counts(agent2, counts_df)

    logging.debug("Counts DataFrame inside count_assignments:")
    logging.debug(counts_df)

    return counts_df


def update_counts(agent, counts_df):
    if agent in counts_df["Name"].values:
        counts_df.loc[counts_df["Name"] == agent, "Assignments"] += 1
    else:
        counts_df = counts_df._append(
            {"Name": agent, "Assignments": 1}, ignore_index=True
        )

    return counts_df


def fill_blank_assignments(schedule_df, counts_df, team_df):
    # Identify cells with blank assignments
    blank_cells = schedule_df.map(lambda x: pd.isnull(x))

    logging.debug("Entering fill_blank_assignments function")
    logging.debug("Blank cells in schedule_df:")

    for index, row in blank_cells.iterrows():
        for col, is_blank in row.items():
            if is_blank:
                logging.debug(
                    f"Index: {index}, Column: {col}, Value: {schedule_df.at[index, col]}"
                )

    for index, row in blank_cells.iterrows():
        logging.debug(f"Filling blank cell(s) at index {index}")

        least_assigned_agent = get_least_assigned_agent(counts_df)
        if least_assigned_agent is not None:
            logging.debug(
                "Assigning duties for least assigned agent:", least_assigned_agent
            )

            # Update counts_df with the assigned agent
            counts_df.loc[counts_df["Name"] == least_assigned_agent, "Assignments"] += 1

            # Update schedule_df with the assigned agent
            for col, is_blank in row.items():
                if is_blank:
                    schedule_df.at[index, col] = least_assigned_agent
                    email_col = f"Email{col[-1]}"  # Adjust the email column based on Agent column
                    schedule_df.at[index, email_col] = team_df.loc[
                        team_df["Name"] == least_assigned_agent, "Email"
                    ].values[0]
        else:
            logging.info("No least assigned agent available. Skipping assignment.")

    logging.debug("Exiting fill_blank_assignments function")

    return schedule_df, counts_df


def get_least_assigned_agent(counts_df):
    # Filter agents with the minimum assignment count
    min_assignments = counts_df["Assignments"].min()
    available_agents = counts_df[counts_df["Assignments"] == min_assignments]

    logging.debug("Available agents:", available_agents)

    if available_agents.empty:
        print("No agents available.")
        return None

    # Choose the agent with the fewest total assignments
    least_assigned_agent = available_agents.loc[
        available_agents["Assignments"].idxmin()
    ]["Name"]

    logging.debug("Selected least assigned agent:", least_assigned_agent)

    return least_assigned_agent


def assign_duties(schedule_df, team_df):
    available_agents = team_df[team_df["Available"] == "yes"]["Name"].tolist()
    random_pairs = generate_random_pairs(available_agents)

    # Use deque for a rotating queue of agents
    agent_queue = deque(random_pairs)

    for index, row in schedule_df.iterrows():
        start_date = datetime.strptime(row["start_date"], "%m-%d-%Y")  # noqa: F841
        end_date = datetime.strptime(row["end_date"], "%m-%d-%Y")  # noqa: F841

        # Pop a pair from the deque
        week_agents = agent_queue.popleft()

        logging.debug(f"Assigning duties for {start_date} to {end_date}")
        logging.debug(f"Chosen agents: {week_agents}")

        schedule_df.at[index, "Agent1"] = week_agents[0]
        schedule_df.at[index, "Email1"] = team_df.loc[
            team_df["Name"] == week_agents[0], "Email"
        ].values[0]

        if week_agents[1]:
            schedule_df.at[index, "Agent2"] = week_agents[1]
            schedule_df.at[index, "Email2"] = team_df.loc[
                team_df["Name"] == week_agents[1], "Email"
            ].values[0]

        # Append the used pair back to the deque
        agent_queue.append(week_agents)

    return schedule_df


def main():
    logging.info("Start script execution")
    print("Start script execution")

    backup_existing_team_list()

    current_employee_data = read_employee_data(team_list_path)
    previous_employee_data = read_employee_data(team_list_path)

    if current_employee_data is not None and validate_team_list(current_employee_data):
        load_and_detect_changes(current_employee_data, previous_employee_data)

        backup_existing_assignments()

        generate_schedule_weeks()

        schedule_df = pd.read_excel(assignment_path, sheet_name="Weeks")
        team_df = pd.read_excel(team_list_path)

        # Assign duties considering availability and balancing
        schedule_df = assign_duties(schedule_df, current_employee_data)

        # Save the updated schedule
        with pd.ExcelWriter(
            assignment_path, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            schedule_df.to_excel(writer, sheet_name="Weeks", index=False)
            counts_df = count_assignments(schedule_df, team_df)
            counts_df.to_excel(writer, sheet_name="Counts", index=False)

        # Check if the number of available agents is odd before filling blank assignments
        available_agents = team_df[team_df["Available"] == "yes"]["Name"].tolist()
        if len(available_agents) % 2 == 1:
            # Fill in blank assignments for an odd number of available agents
            schedule_df, counts_df = fill_blank_assignments(
                schedule_df, counts_df, team_df
            )

        # Save the final schedule
        with pd.ExcelWriter(
            assignment_path, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            schedule_df.to_excel(writer, sheet_name="Weeks", index=False)
            counts_df.to_excel(writer, sheet_name="Counts", index=False)

        logging.info("Script execution completed.\n")
        print("Script execution completed.\n")

    else:
        error_message = "Problem reading from team_list.xlsx"
        logging.error(f"{error_message}\n")
        print(f"{error_message}\n")


if __name__ == "__main__":
    main()
