# Time Tracker Application

A simple application for tracking time spent on different opportunities and activities.

## Features

- Track time spent on different projects/opportunities
- Categorize time by activity type
- Export time logs as CSV or Excel files
- Generate daily and weekly reports
- Save data persistently between sessions
- Upload and use your own CRM data from Excel files

## CRM Data Format

The application supports loading CRM data from Excel files. Your Excel file should have the following columns:
- `Record Id` - Unique identifier for each opportunity
- `Deal Name` - Name of the opportunity/deal
- `Company Name (Company Name)` - Name of the company
- `Deal Owner` - Name of the deal owner

A sample file is included with the application, which you can use as a template.

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

### Building for Multiple Platforms

To build for Windows:
- You must run the build script on a Windows machine
- Transfer the project files to a Windows computer and run `build.py` there

To build for macOS:
- You must run the build script on a macOS machine

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

1. On first launch, you'll be prompted to select a CRM data file or use the sample data
2. Select your role, activity type, and opportunity from the dropdowns
3. Click "Start" to begin tracking time
4. Use "Pause" to temporarily stop the timer and "Resume" to continue
5. Click "Stop" when you've completed the activity
6. Export your time logs using the export options (daily or weekly, CSV or Excel)
7. If needed, you can change the CRM data file using the "Change CRM Data File" button

## Data Storage

The application stores time logs and your selected CRM data file in the following location:
- Windows: `C:\Users\<username>\AppData\Roaming\TimeTracker\`
- macOS/Linux: `~/.timetracker/` 