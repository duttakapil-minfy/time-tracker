# Time Tracker Application

A simple application for tracking time spent on different opportunities and activities.

## Features

- Track time spent on different projects/opportunities
- Categorize time by activity type
- Export time logs as CSV or Excel files
- Generate daily and weekly reports
- Save data persistently between sessions

## Building the Application

### Prerequisites

- Python 3.7 or higher
- Required Python packages (will be installed automatically during build):
  - pandas
  - openpyxl
  - pyinstaller

### Building the Executable

1. Clone or download this repository to your local machine
2. Open a terminal/command prompt and navigate to the project directory
3. Run the build script:

   ```
   python build.py
   ```

4. The executable will be created in the `dist` directory:
   - Windows: `dist/TimeTracker.exe`
   - macOS: `dist/TimeTracker.app`
   - Linux: `dist/TimeTracker`

## Running the Application

### From Source

If you prefer to run the application from source:

1. Ensure you have Python 3.7+ installed
2. Install required packages:
   ```
   pip install pandas openpyxl
   ```
3. Run the application:
   ```
   python time-tracker.py
   ```

### From Executable

Simply double-click the executable file in the `dist` directory:
- Windows: `TimeTracker.exe`
- macOS: `TimeTracker.app` (You may need to right-click and select 'Open' the first time)
- Linux: `TimeTracker`

## Usage

1. Select your role, activity type, and opportunity from the dropdowns
2. Click "Start" to begin tracking time
3. Use "Pause" to temporarily stop the timer and "Resume" to continue
4. Click "Stop" when you've completed the activity
5. Export your time logs using the export options (daily or weekly, CSV or Excel)

## Data Storage

The application stores time logs in the following location:
- Windows: `C:\Users\<username>\AppData\Roaming\TimeTracker\time_logs_data.json`
- macOS/Linux: `~/.timetracker/time_logs_data.json` 