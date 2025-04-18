import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime, timedelta
import time
import csv
import os
import json
import calendar
import shutil
from collections import defaultdict
from resource_path import resource_path

class TimeTrackerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Opportunity Time Tracker")
        
        # Set application icon (for Windows executable)
        try:
            self.root.iconbitmap(resource_path("icon.ico"))
        except:
            pass  # Silently fail if icon doesn't exist
        
        # Initialize time logs storage
        self.app_data_dir = self.get_app_data_dir()
        self.time_logs_file = os.path.join(self.app_data_dir, "time_logs_data.json")
        self.crm_data_file = os.path.join(self.app_data_dir, "crm_data.xlsx")
        self.time_logs = self.load_time_logs()
        
        # Define roles and activities
        self.roles = ["Pre-Sales"]  # Default role for now, can be expanded later
        
        # Activity types based on the image
        self.activities = [
            "Solution Innovation and Improvement",
            "Client and Partner Engagement",
            "Solution Design and Architecture",
            "Proposal Support",
            "Solution Documentation",
            "Internal Meetings",
            "Training and Development",
            "Administrative Tasks"
        ]
        
        self.current_timer = None
        self.start_time = None
        self.running = False
        self.df = None
        
        # GUI Setup - Create main window structure first
        self.setup_main_window()
        
        # Load CRM data or prompt user to select a file
        self.load_crm_data()
        
    def setup_main_window(self):
        """Set up the main window structure"""
        # Main content frame
        self.main_frame = tk.Frame(self.root, padx=10, pady=10)
        self.main_frame.pack(fill="both", expand=True)
        
        # Create a placeholder for the main interface
        # The actual widgets will be created after we have CRM data
        self.placeholder_frame = tk.Frame(self.main_frame)
        self.placeholder_frame.pack(fill="both", expand=True)
        
    def get_app_data_dir(self):
        """Get or create application data directory"""
        # Create an app-specific data directory in user's home
        if os.name == 'nt':  # Windows
            app_data = os.path.join(os.environ['APPDATA'], 'TimeTracker')
        else:  # macOS and Linux
            app_data = os.path.join(os.path.expanduser('~'), '.timetracker')
            
        # Create directory if it doesn't exist
        if not os.path.exists(app_data):
            os.makedirs(app_data)
            
        return app_data
        
    def load_time_logs(self):
        """Load existing time logs from JSON file"""
        if os.path.exists(self.time_logs_file):
            try:
                with open(self.time_logs_file, 'r') as f:
                    return json.load(f)
            except Exception as e:
                messagebox.showwarning("Warning", f"Failed to load existing time logs: {str(e)}")
        return []
        
    def save_time_logs(self):
        """Save time logs to JSON file for persistence"""
        try:
            with open(self.time_logs_file, 'w') as f:
                json.dump(self.time_logs, f)
        except Exception as e:
            messagebox.showwarning("Warning", f"Failed to save time logs: {str(e)}")

    def load_crm_data(self):
        """Load CRM data from file or prompt user to select a file"""
        # Check if we have a saved CRM data file
        if os.path.exists(self.crm_data_file):
            try:
                self.df = pd.read_excel(self.crm_data_file)
                if len(self.df) > 0:
                    self.create_widgets()
                    return
            except Exception as e:
                # If there's an error, we'll prompt the user for a new file
                pass
        
        # We didn't find a valid saved file, so ask the user to upload one
        self.prompt_for_crm_file()
    
    def prompt_for_crm_file(self):
        """Show interface for loading a CRM data file"""
        # Clear placeholder
        for widget in self.placeholder_frame.winfo_children():
            widget.destroy()
        
        # Create labels and buttons for file selection
        tk.Label(
            self.placeholder_frame, 
            text="Welcome to Time Tracker",
            font=("Arial", 16)
        ).pack(pady=20)
        
        tk.Label(
            self.placeholder_frame,
            text="No CRM data file found. Please select your CRM data Excel file to continue.",
            wraplength=400
        ).pack(pady=10)
        
        tk.Label(
            self.placeholder_frame,
            text="The file should have columns for 'Record Id', 'Deal Name', 'Company Name (Company Name)', and 'Deal Owner'.",
            wraplength=400,
            font=("Arial", 9),
            fg="gray"
        ).pack(pady=10)
        
        # Try to load a default file from the resources if it exists (for first-time use)
        default_file_btn = tk.Button(
            self.placeholder_frame,
            text="Use Sample Data",
            command=self.use_sample_data
        )
        default_file_btn.pack(pady=5)
        
        # Or select a custom file
        select_file_btn = tk.Button(
            self.placeholder_frame,
            text="Select CRM Data File",
            command=self.select_crm_file
        )
        select_file_btn.pack(pady=20)
    
    def use_sample_data(self):
        """Use sample data from resources (if available)"""
        try:
            sample_file = resource_path("Kapil-Dutta-Open-Deals.xlsx")
            if os.path.exists(sample_file):
                # Copy the sample file to the app data directory
                shutil.copy(sample_file, self.crm_data_file)
                self.df = pd.read_excel(self.crm_data_file)
                self.create_widgets()
            else:
                messagebox.showerror("Error", "Sample data file not found. Please select your own file.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load sample data: {str(e)}")
    
    def select_crm_file(self):
        """Allow user to select a CRM data file"""
        file_path = filedialog.askopenfilename(
            title="Select CRM Data File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                # Try to load the selected file
                temp_df = pd.read_excel(file_path)
                
                # Validate that it has the required columns
                required_columns = ['Record Id', 'Deal Name', 'Company Name (Company Name)', 'Deal Owner']
                missing_columns = [col for col in required_columns if col not in temp_df.columns]
                
                if missing_columns:
                    messagebox.showerror("Invalid File", 
                                        f"The selected file is missing the following required columns: {', '.join(missing_columns)}\n\n"
                                        "Please select a file with the correct format.")
                    return
                
                # Copy to app data directory and use it
                shutil.copy(file_path, self.crm_data_file)
                self.df = temp_df
                self.create_widgets()
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load the selected file: {str(e)}")
    
    def create_widgets(self):
        """Create the main application widgets after loading CRM data"""
        # Clear placeholder frame
        for widget in self.placeholder_frame.winfo_children():
            widget.destroy()
        
        # ========== Selection Section ==========
        selection_frame = tk.LabelFrame(self.placeholder_frame, text="Project Selection")
        selection_frame.pack(fill="x", padx=5, pady=5)
        
        # Role Selection
        tk.Label(selection_frame, text="Role:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.role_combo = ttk.Combobox(selection_frame, values=self.roles, width=20)
        self.role_combo.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.role_combo.set("Pre-Sales")  # Default role
        
        # CRM Data File Button - Allow changing the data source
        change_file_btn = tk.Button(
            selection_frame, 
            text="Change CRM Data File", 
            command=self.select_crm_file
        )
        change_file_btn.grid(row=0, column=2, padx=5, pady=5, sticky="e")
        
        # Activity Selection
        tk.Label(selection_frame, text="Activity:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.activity_combo = ttk.Combobox(selection_frame, values=self.activities, width=40)
        self.activity_combo.grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky="w")
        
        # Opportunity Selection
        tk.Label(selection_frame, text="Opportunity:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.opportunity_combo = ttk.Combobox(selection_frame, 
                                            values=[f"{row['Deal Name']} ({row['Record Id']})" for _, row in self.df.iterrows()],
                                            width=40)
        self.opportunity_combo.grid(row=2, column=1, columnspan=2, padx=5, pady=5, sticky="w")
        
        # Optional Comment Field
        tk.Label(selection_frame, text="Comments:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.comment_entry = tk.Entry(selection_frame, width=40)
        self.comment_entry.grid(row=3, column=1, columnspan=2, padx=5, pady=5, sticky="w")
        
        # ========== Timer Section ==========
        timer_frame = tk.LabelFrame(self.placeholder_frame, text="Time Tracking")
        timer_frame.pack(fill="x", padx=5, pady=10)
        
        # Timer Display
        self.timer_label = tk.Label(timer_frame, text="00:00:00", font=('Arial', 20))
        self.timer_label.pack(pady=10)
        
        # Timer Control Buttons
        timer_control_frame = tk.Frame(timer_frame)
        timer_control_frame.pack(pady=5)
        
        self.start_btn = tk.Button(timer_control_frame, text="Start", command=self.start_timer)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        
        self.pause_btn = tk.Button(timer_control_frame, text="Pause", command=self.pause_timer, state='disabled')
        self.pause_btn.pack(side=tk.LEFT, padx=5)
        
        self.stop_btn = tk.Button(timer_control_frame, text="Stop", command=self.stop_timer, state='disabled')
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        
        # ========== Export Section ==========
        export_frame = tk.LabelFrame(self.placeholder_frame, text="Export Options")
        export_frame.pack(fill="x", padx=5, pady=5)
        
        # Daily export options
        daily_frame = tk.Frame(export_frame)
        daily_frame.pack(fill="x", pady=5)
        
        tk.Label(daily_frame, text="Daily Export:").pack(side=tk.LEFT, padx=5)
        self.export_daily_csv_btn = tk.Button(daily_frame, text="CSV", 
                                            command=lambda: self.export_logs("csv", "daily"))
        self.export_daily_csv_btn.pack(side=tk.LEFT, padx=5)
        
        self.export_daily_excel_btn = tk.Button(daily_frame, text="Excel", 
                                               command=lambda: self.export_logs("excel", "daily"))
        self.export_daily_excel_btn.pack(side=tk.LEFT, padx=5)
        
        # Weekly export options
        weekly_frame = tk.Frame(export_frame)
        weekly_frame.pack(fill="x", pady=5)
        
        tk.Label(weekly_frame, text="Weekly Export:").pack(side=tk.LEFT, padx=5)
        self.export_weekly_csv_btn = tk.Button(weekly_frame, text="CSV", 
                                              command=lambda: self.export_logs("csv", "weekly"))
        self.export_weekly_csv_btn.pack(side=tk.LEFT, padx=5)
        
        self.export_weekly_excel_btn = tk.Button(weekly_frame, text="Excel", 
                                                command=lambda: self.export_logs("excel", "weekly"))
        self.export_weekly_excel_btn.pack(side=tk.LEFT, padx=5)
        
        # Show current logs count
        self.logs_info_label = tk.Label(self.placeholder_frame, text=f"Saved logs: {len(self.time_logs)}")
        self.logs_info_label.pack(pady=5)

    def start_timer(self):
        if not self.running:
            # Validate required fields
            if not self.opportunity_combo.get():
                messagebox.showwarning("Warning", "Please select an opportunity!")
                return
            if not self.activity_combo.get():
                messagebox.showwarning("Warning", "Please select an activity type!")
                return
            
            self.running = True
            self.start_time = time.time()
            self.start_btn.config(state='disabled')
            self.pause_btn.config(state='normal')
            self.stop_btn.config(state='normal')
            self.update_timer()

    def pause_timer(self):
        if self.running:
            self.running = False
            self.root.after_cancel(self.current_timer)
            self.start_btn.config(state='normal', text='Resume')
            self.pause_btn.config(state='disabled')

    def stop_timer(self):
        if self.start_time:
            end_time = time.time()
            self.running = False
            self.root.after_cancel(self.current_timer)
            
            # Get opportunity details
            selected_opp = self.opportunity_combo.get()
            record_id = selected_opp.split('(')[-1].strip(')')
            deal_name = selected_opp.split('(')[0].strip()
            
            # Find the corresponding row in the DataFrame
            row = self.df[self.df['Record Id'] == record_id]
            if not row.empty:
                company = row['Company Name (Company Name)'].iloc[0]
                owner = row['Deal Owner'].iloc[0]
            else:
                company = "Unknown"
                owner = "Unknown"
            
            current_time = datetime.now()
            
            # Get role and activity
            role = self.role_combo.get()
            activity = self.activity_combo.get()
            comment = self.comment_entry.get()
            
            self.time_logs.append({
                'Date': current_time.strftime('%Y-%m-%d'),
                'Timestamp': current_time.timestamp(),  # Store timestamp for accurate weekly filtering
                'Record Id': record_id,
                'Deal Name': deal_name,
                'Company Name': company,
                'Deal Owner': owner,
                'Role': role,
                'Activity': activity,
                'Comment': comment,
                'Start Time': datetime.fromtimestamp(self.start_time).strftime('%H:%M:%S'),
                'End Time': datetime.fromtimestamp(end_time).strftime('%H:%M:%S'),
                'Duration (seconds)': int(end_time - self.start_time)
            })
            
            # Save time logs to file for persistence
            self.save_time_logs()
            
            # Update logs info label
            self.logs_info_label.config(text=f"Saved logs: {len(self.time_logs)}")
            
            # Reset the timer and buttons
            self.timer_label.config(text="00:00:00")
            self.start_btn.config(state='normal', text='Start')
            self.pause_btn.config(state='disabled')
            self.stop_btn.config(state='disabled')
            self.start_time = None
            
            # Clear the comment field for next entry
            self.comment_entry.delete(0, tk.END)

    def update_timer(self):
        if self.running:
            elapsed = int(time.time() - self.start_time)
            hours = elapsed // 3600
            minutes = (elapsed % 3600) // 60
            seconds = elapsed % 60
            self.timer_label.config(text=f"{hours:02d}:{minutes:02d}:{seconds:02d}")
            self.current_timer = self.root.after(1000, self.update_timer)

    def export_logs(self, format_type="csv", period="daily"):
        if not self.time_logs:
            messagebox.showinfo("Info", "No time logs to export!")
            return
        
        # Filter logs based on period
        logs_to_export = []
        today = datetime.now().date()
        
        if period == "daily":
            # Filter logs for today
            logs_to_export = [log for log in self.time_logs 
                             if datetime.strptime(log['Date'], '%Y-%m-%d').date() == today]
            period_str = f"daily_{today.strftime('%Y-%m-%d')}"
            if not logs_to_export:
                messagebox.showinfo("Info", "No time logs for today!")
                return
                
        elif period == "weekly":
            # Calculate the start of the current week (Monday)
            start_of_week = today - timedelta(days=today.weekday())
            # Calculate the end of the current week (Sunday)
            end_of_week = start_of_week + timedelta(days=6)
            # Filter logs for this week
            logs_to_export = [log for log in self.time_logs 
                             if datetime.strptime(log['Date'], '%Y-%m-%d').date() >= start_of_week and
                                datetime.strptime(log['Date'], '%Y-%m-%d').date() <= end_of_week]
            period_str = f"weekly_{start_of_week.strftime('%Y-%m-%d')}_to_{end_of_week.strftime('%Y-%m-%d')}"
            if not logs_to_export:
                messagebox.showinfo("Info", "No time logs for this week!")
                return
        
        # Create a DataFrame from filtered time logs
        logs_df = pd.DataFrame(logs_to_export)
        
        # Add formatted duration (HH:MM:SS)
        if 'Duration (seconds)' in logs_df.columns:
            logs_df['Duration (HH:MM:SS)'] = logs_df['Duration (seconds)'].apply(
                lambda x: f"{x//3600:02d}:{(x%3600)//60:02d}:{x%60:02d}")
        
        # For cleaner output, drop timestamp column which is only for internal use
        if 'Timestamp' in logs_df.columns:
            logs_df = logs_df.drop(columns=['Timestamp'])
        
        # Choose file path for export
        if format_type == "csv":
            default_filename = f"time_logs_{period_str}.csv"
            filetypes = [("CSV files", "*.csv"), ("All files", "*.*")]
            file_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=filetypes,
                initialfile=default_filename
            )
            
            if file_path:
                # Export detailed logs as CSV
                logs_df.to_csv(file_path, index=False)
                
                # Create weekly summary timesheet-like format
                self.export_weekly_summary(logs_to_export, file_path.replace('.csv', '_weekly_summary.csv'), "csv")
                
                messagebox.showinfo("Success", 
                                   f"Time logs exported to {os.path.basename(file_path)}\n"
                                   f"Weekly summary exported to {os.path.basename(file_path.replace('.csv', '_weekly_summary.csv'))}")
                
        elif format_type == "excel":
            default_filename = f"time_logs_{period_str}.xlsx"
            filetypes = [("Excel files", "*.xlsx"), ("All files", "*.*")]
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=filetypes,
                initialfile=default_filename
            )
            
            if file_path:
                # Export as Excel with multiple sheets
                with pd.ExcelWriter(file_path) as writer:
                    logs_df.to_excel(writer, sheet_name='Detailed Logs', index=False)
                    
                    # Add weekly timesheet-like format
                    self.export_weekly_summary(logs_to_export, writer, "excel")
                
                messagebox.showinfo("Success", 
                                   f"Time logs exported to {os.path.basename(file_path)}\n"
                                   f"(Includes both detailed logs and weekly summary)")
    
    def export_weekly_summary(self, logs, output, format_type):
        """Create a weekly summary similar to the timesheet image"""
        if not logs:
            return
        
        # Get the date range for the week
        dates = [datetime.strptime(log['Date'], '%Y-%m-%d').date() for log in logs]
        min_date = min(dates)
        max_date = max(dates)
        
        # Create a date range for the full week (Monday to Sunday)
        start_of_week = min_date - timedelta(days=min_date.weekday())
        end_of_week = start_of_week + timedelta(days=6)
        date_range = [start_of_week + timedelta(days=i) for i in range(7)]
        
        # Create column headers with weekday names and dates
        weekdays = ['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN']
        # Get the actual month abbreviation from the date range
        date_headers = [f"{d.strftime('%b').upper()} {d.day}" for d in date_range]
        
        # Group logs by activity and date
        activity_time = defaultdict(lambda: defaultdict(int))
        for log in logs:
            activity = log['Activity']
            log_date = datetime.strptime(log['Date'], '%Y-%m-%d').date()
            weekday_idx = (log_date - start_of_week).days
            if 0 <= weekday_idx < 7:  # Ensure it's within our week
                duration_seconds = log['Duration (seconds)']
                activity_time[activity][weekday_idx] += duration_seconds
        
        # Create a list to hold the data for the DataFrame
        data = []
        
        # Aggregate attendance hours (total per day)
        daily_totals = [0] * 7
        for activity, day_durations in activity_time.items():
            for day_idx, duration in day_durations.items():
                daily_totals[day_idx] += duration
        
        # Convert seconds to hours:minutes format
        attendance_row = ['ATTENDANCE HOURS']
        for day_total in daily_totals:
            hours = day_total // 3600
            minutes = (day_total % 3600) // 60
            attendance_row.append(f"{hours}h {minutes}m" if day_total > 0 else "0h 0m")
        
        # Add placeholder for TASK TOTAL column
        attendance_row.append("")  # Task total for attendance
        data.append(attendance_row)
        
        # Add a blank row
        data.append([""] * 9)
        
        # Add activity rows with time entries
        task_totals = {}
        
        for activity in sorted(activity_time.keys()):
            row = [activity]
            total_activity_seconds = 0
            
            for day_idx in range(7):
                duration_seconds = activity_time[activity].get(day_idx, 0)
                total_activity_seconds += duration_seconds
                
                if duration_seconds > 0:
                    hours = duration_seconds // 3600
                    minutes = (duration_seconds % 3600) // 60
                    row.append(f"{hours}:{minutes:02d}")
                else:
                    row.append("0:00")
            
            # Calculate task total (HH:MM)
            hours = total_activity_seconds // 3600
            minutes = (total_activity_seconds % 3600) // 60
            task_total = f"{hours}:{minutes:02d}"
            row.append(task_total)
            
            data.append(row)
            
            # Add comment rows - populate from comments if available
            comment_row = ["COMMENT"] + [""] * 8
            data.append(comment_row)
        
        # Calculate weekly total in seconds
        weekly_total_seconds = sum(daily_totals)
        hours = weekly_total_seconds // 3600
        minutes = (weekly_total_seconds % 3600) // 60
        weekly_total = f"{hours}:{minutes:02d}"
        
        # Add a row for totals
        totals_row = ["Total hours/day"]
        for day_total in daily_totals:
            hours = day_total // 3600
            minutes = (day_total % 3600) // 60
            totals_row.append(f"{hours}:{minutes:02d}")
        totals_row.append(weekly_total)
        data.append(totals_row)
        
        # Create column headers
        columns = ['PROJECTS'] + date_headers + ['TASK TOTAL\nHRS/WEEK']
        
        # Create the DataFrame
        summary_df = pd.DataFrame(data, columns=columns)
        
        # Export the summary
        if format_type == "csv":
            summary_df.to_csv(output, index=False)
        else:  # Excel
            # Adjust column widths for better display
            summary_df.to_excel(output, sheet_name='Weekly Timesheet', index=False)
            worksheet = output.sheets['Weekly Timesheet']
            
            # In a complete implementation, you would add Excel formatting here
            # such as column widths, cell styles, etc.

if __name__ == "__main__":
    root = tk.Tk()
    app = TimeTrackerApp(root)
    root.mainloop()