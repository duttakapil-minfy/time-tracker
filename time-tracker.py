import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime, timedelta
import time
import csv
import os
import json
import calendar
from collections import defaultdict

class TimeTrackerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Opportunity Time Tracker")
        
        # Load opportunities from Excel file
        try:
            excel_file = "Kapil-Dutta-Open-Deals.xlsx"
            self.df = pd.read_excel(excel_file)
            if len(self.df) == 0:
                messagebox.showerror("Error", f"No data found in {excel_file}")
                self.root.destroy()
                return
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load data: {str(e)}")
            self.root.destroy()
            return
        
        # Initialize time logs storage
        self.time_logs_file = "time_logs_data.json"
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
        
        # GUI Setup
        self.create_widgets()
        
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

    def create_widgets(self):
        # Main content frame
        main_frame = tk.Frame(self.root, padx=10, pady=10)
        main_frame.pack(fill="both", expand=True)
        
        # ========== Selection Section ==========
        selection_frame = tk.LabelFrame(main_frame, text="Project Selection")
        selection_frame.pack(fill="x", padx=5, pady=5)
        
        # Role Selection
        tk.Label(selection_frame, text="Role:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.role_combo = ttk.Combobox(selection_frame, values=self.roles, width=20)
        self.role_combo.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.role_combo.set("Pre-Sales")  # Default role
        
        # Activity Selection
        tk.Label(selection_frame, text="Activity:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.activity_combo = ttk.Combobox(selection_frame, values=self.activities, width=40)
        self.activity_combo.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        
        # Opportunity Selection
        tk.Label(selection_frame, text="Opportunity:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.opportunity_combo = ttk.Combobox(selection_frame, 
                                            values=[f"{row['Deal Name']} ({row['Record Id']})" for _, row in self.df.iterrows()],
                                            width=40)
        self.opportunity_combo.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        
        # Optional Comment Field
        tk.Label(selection_frame, text="Comments:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.comment_entry = tk.Entry(selection_frame, width=40)
        self.comment_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        
        # ========== Timer Section ==========
        timer_frame = tk.LabelFrame(main_frame, text="Time Tracking")
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
        export_frame = tk.LabelFrame(main_frame, text="Export Options")
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
        self.logs_info_label = tk.Label(main_frame, text=f"Saved logs: {len(self.time_logs)}")
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
        date_headers = [f"MAR {d.day}" for d in date_range]  # Using month from the image
        # In real implementation, you would use:
        # date_headers = [f"{d.strftime('%b').upper()} {d.day}" for d in date_range]
        
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
            
            # Add comment rows (blank in this implementation, could be populated from log comments)
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