# Scheduler

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
# For Windows (command line with Admin)
venv\Scripts\activate.ps1

# For macOS/Linux
source venv/bin/activate
```

### Install Dependencies

Install the project dependencies listed in the `requirements.txt` file:

```bash
pip install -r requirements.txt
```

(Reads the `requirements.txt` file and installs all the specified packages.)

### Deactivate the Virtual Environment (Optional)

If you used a virtual environment, you can deactivate it when you're done working on the project:

```bash
deactivate
```


# Documentation