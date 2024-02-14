# Schedule Generator

 - Automates the process of generating the on-point schedule based on agent availability.

## Description

- **Backup:** Before making any changes, the script backs up both the team list and existing assignments to the `history` directory.
- **Logging:** Events and errors are logged to a file (`logging/script_events.log`) for debugging.
- **Availability Check:** The team's availability is validated through the "Available" column in the `team_list.xlsx` file.
- **Dynamic Scheduling:** The script generates a schedule for a specified number of weeks (default is 52 weeks), ensuring coverage and balancing workloads.
- **Random Pairing:** Duties are assigned by pairing employees randomly, maintaining a fair distribution.
- **Fill Blank Assignments:** In case of an odd number of available agents, the script fills blank assignments to ensure all agents have similar workloads.

## Usage

1. **Install Dependencies:**

    - See "Installation" section below for Virtual Environment setup

2. **Setup:**
   - Folder structure will be created if it doesn't exist.
   - Ensure your team list is in the `data` directory as `team_list.xlsx`.
   - `team_list.xlsx` must contain 3 columns: Name, Email, and Available (case sensitive).
   - The 'Available' column must contain either yes or no values (not case sensitive).

3. **Review Output:**
   - The script will generate a schedule in the `data` directory named `assignments.xlsx`.
   - Logs will be available in the `logging` directory.

## Configuration

- Adjust constants in the script:
  - `NUM_WEEKS`: Number of weeks for the schedule.
  - `log_directory`, `data_directory`, `history_directory`: Directories for logs, data, and backup history.

## Logging

- Logs are stored in `logging/script_events.log`.
- Information, errors, and changes in team composition are logged.

## Installation

To set up the project and install the required Python libraries, follow the steps below.
(Assumes Python 3 is installed)

### Clone the Repo

### Create a Virtual Environment (Optional but recommended)

Isolates dependencies and avoids globally installing the required libraries.
If you don't have `virtualenv` installed, you can install it using:

```bash
pip install virtualenv
```

Create a virtual environment:

```bash
python -m venv venv  # For Windows
# OR
python3 -m venv venv  # For macOS/Linux
```

Activate the virtual environment:

```bash
venv\Scripts\activate.ps1 # For Windows (command line with Admin)
# OR
source venv/bin/activate # For macOS/Linux
```

### Install Dependencies

Install the project dependencies listed in the `requirements.txt` file:

```bash
pip install -r requirements.txt
```

(Reads the `requirements.txt` file and installs all the specified packages.)

### Execute the script:

```bash
python main.py  # For Windows
# OR
python3 main.py  # For macOS/Linux
```