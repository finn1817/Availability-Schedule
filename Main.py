import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import pickle

class Workplace:
    def __init__(self, name, hours_of_operation=None, workers=None, shifts=None):
        self.name = name
        self.hours_of_operation = hours_of_operation or {"Monday": ("9:00", "17:00"), 
                                                         "Tuesday": ("9:00", "17:00"),
                                                         "Wednesday": ("9:00", "17:00"),
                                                         "Thursday": ("9:00", "17:00"),
                                                         "Friday": ("9:00", "17:00"),
                                                         "Saturday": ("10:00", "16:00"),
                                                         "Sunday": ("10:00", "16:00")}
        self.workers = workers or []
        self.shifts = shifts or {}
        self.excel_file = None
    
    def __str__(self):
        return self.name

class WorkplaceSchedulerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Workplace Scheduler")
        self.root.geometry("900x600")
        
        # State variables
        self.workplaces = []
        self.current_workplace = None
        
        # Create main frames
        self.create_main_frame()
        self.load_workplaces()
        
        # Start with home screen
        self.show_home_screen()
    
    def create_main_frame(self):
        # Main container
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header frame
        self.header_frame = ttk.Frame(self.main_frame)
        self.header_frame.pack(fill=tk.X, pady=10)
        
        self.title_label = ttk.Label(self.header_frame, text="Workplace Scheduler", font=("Arial", 18, "bold"))
        self.title_label.pack(side=tk.LEFT)
        
        self.home_button = ttk.Button(self.header_frame, text="Home", command=self.show_home_screen)
        self.home_button.pack(side=tk.RIGHT)
        
        # Content frame - will be cleared and repopulated based on current view
        self.content_frame = ttk.Frame(self.main_frame)
        self.content_frame.pack(fill=tk.BOTH, expand=True, pady=10)
    
    def clear_content_frame(self):
        for widget in self.content_frame.winfo_children():
            widget.destroy()
    
    def show_home_screen(self):
        self.clear_content_frame()
        self.title_label.config(text="Workplace Manager")
        
        # Workplaces list frame
        list_frame = ttk.LabelFrame(self.content_frame, text="Your Workplaces")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Scrollable frame for workplaces
        scroll_frame = ttk.Frame(list_frame)
        scroll_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        scrollbar = ttk.Scrollbar(scroll_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.workplace_listbox = tk.Listbox(scroll_frame, height=10, width=50, 
                                           font=("Arial", 12), 
                                           selectmode=tk.SINGLE,
                                           yscrollcommand=scrollbar.set)
        self.workplace_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.workplace_listbox.yview)
        
        # Populate workplace list
        for workplace in self.workplaces:
            self.workplace_listbox.insert(tk.END, workplace.name)
        
        # Buttons frame
        buttons_frame = ttk.Frame(list_frame)
        buttons_frame.pack(fill=tk.X, pady=10)
        
        add_button = ttk.Button(buttons_frame, text="Add Workplace", command=self.add_workplace)
        add_button.pack(side=tk.LEFT, padx=5)
        
        select_button = ttk.Button(buttons_frame, text="Select Workplace", command=self.select_workplace)
        select_button.pack(side=tk.LEFT, padx=5)
        
        remove_button = ttk.Button(buttons_frame, text="Remove Workplace", command=self.remove_workplace)
        remove_button.pack(side=tk.LEFT, padx=5)
        
        save_button = ttk.Button(buttons_frame, text="Save All Workplaces", command=self.save_workplaces)
        save_button.pack(side=tk.RIGHT, padx=5)
    
    def add_workplace(self):
        # Dialog to add new workplace
        add_window = tk.Toplevel(self.root)
        add_window.title("Add New Workplace")
        add_window.geometry("400x200")
        add_window.grab_set()  # Modal window
        
        ttk.Label(add_window, text="Workplace Name:").pack(pady=10)
        name_entry = ttk.Entry(add_window, width=40)
        name_entry.pack(pady=10)
        name_entry.focus()
        
        def save_new_workplace():
            name = name_entry.get().strip()
            if name:
                # Check for duplicates
                if any(wp.name == name for wp in self.workplaces):
                    messagebox.showerror("Error", f"Workplace '{name}' already exists!")
                    return
                
                # Create new workplace and add to list
                new_workplace = Workplace(name)
                self.workplaces.append(new_workplace)
                self.workplace_listbox.insert(tk.END, name)
                add_window.destroy()
                messagebox.showinfo("Success", f"Workplace '{name}' added successfully!")
            else:
                messagebox.showerror("Error", "Please enter a workplace name!")
        
        ttk.Button(add_window, text="Save", command=save_new_workplace).pack(pady=20)
    
    def select_workplace(self):
        selection = self.workplace_listbox.curselection()
        if not selection:
            messagebox.showinfo("Info", "Please select a workplace first!")
            return
        
        index = selection[0]
        self.current_workplace = self.workplaces[index]
        self.show_workplace_screen()
    
    def remove_workplace(self):
        selection = self.workplace_listbox.curselection()
        if not selection:
            messagebox.showinfo("Info", "Please select a workplace to remove!")
            return
        
        index = selection[0]
        workplace = self.workplaces[index]
        
        # Confirm deletion
        confirm = messagebox.askyesno("Confirm Deletion", 
                                     f"Are you sure you want to delete '{workplace.name}'?\nThis action cannot be undone.")
        if confirm:
            self.workplaces.pop(index)
            self.workplace_listbox.delete(index)
            messagebox.showinfo("Success", f"Workplace '{workplace.name}' removed successfully!")
    
    def show_workplace_screen(self):
        self.clear_content_frame()
        self.title_label.config(text=f"Workplace: {self.current_workplace.name}")
        
        # Tabs for different workplace settings
        notebook = ttk.Notebook(self.content_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # Data tab
        data_frame = ttk.Frame(notebook)
        notebook.add(data_frame, text="Data Import")
        
        ttk.Label(data_frame, text="Import worker data from Excel:", font=("Arial", 12)).pack(pady=10)
        
        file_frame = ttk.Frame(data_frame)
        file_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.file_path_var = tk.StringVar()
        if self.current_workplace.excel_file:
            self.file_path_var.set(self.current_workplace.excel_file)
        
        file_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, width=50)
        file_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        def browse_file():
            filename = filedialog.askopenfilename(
                title="Select Excel File",
                filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
            )
            if filename:
                self.file_path_var.set(filename)
        
        browse_button = ttk.Button(file_frame, text="Browse...", command=browse_file)
        browse_button.pack(side=tk.RIGHT, padx=5)
        
        def import_data():
            file_path = self.file_path_var.get()
            if not file_path or not os.path.exists(file_path):
                messagebox.showerror("Error", "Please select a valid Excel file!")
                return
            
            try:
                # Read Excel file
                df = pd.read_excel(file_path)
                
                # Basic validation
                required_columns = ["Name", "Position", "Availability"]
                missing_columns = [col for col in required_columns if col not in df.columns]
                
                if missing_columns:
                    messagebox.showerror("Error", f"Missing required columns: {', '.join(missing_columns)}")
                    return
                
                # Process and store data
                self.current_workplace.workers = df.to_dict('records')
                self.current_workplace.excel_file = file_path
                
                # Display preview
                self.preview_text.delete(1.0, tk.END)
                self.preview_text.insert(tk.END, f"Successfully imported {len(df)} workers.\n\n")
                self.preview_text.insert(tk.END, "Preview of data:\n")
                self.preview_text.insert(tk.END, df.head().to_string())
                
                messagebox.showinfo("Success", f"Successfully imported {len(df)} workers from {file_path}")
                
            except Exception as e:
                messagebox.showerror("Import Error", f"Error importing data: {str(e)}")
        
        import_button = ttk.Button(data_frame, text="Import Data", command=import_data)
        import_button.pack(pady=10)
        
        preview_frame = ttk.LabelFrame(data_frame, text="Data Preview")
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Add scrollbar for preview
        preview_scroll = ttk.Scrollbar(preview_frame)
        preview_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.preview_text = tk.Text(preview_frame, wrap=tk.WORD, height=15, 
                                   yscrollcommand=preview_scroll.set)
        self.preview_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        preview_scroll.config(command=self.preview_text.yview)
        
        # Show existing data if available
        if self.current_workplace.workers:
            self.preview_text.insert(tk.END, f"{len(self.current_workplace.workers)} workers loaded.\n\n")
            if len(self.current_workplace.workers) > 0:
                df = pd.DataFrame(self.current_workplace.workers)
                self.preview_text.insert(tk.END, "Preview of data:\n")
                self.preview_text.insert(tk.END, df.head().to_string())
        
        # Hours tab
        hours_frame = ttk.Frame(notebook)
        notebook.add(hours_frame, text="Hours of Operation")
        
        ttk.Label(hours_frame, text="Set hours of operation for each day:", 
                 font=("Arial", 12)).pack(pady=10)
        
        days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        hours_entries = {}
        
        for day in days:
            day_frame = ttk.Frame(hours_frame)
            day_frame.pack(fill=tk.X, padx=20, pady=5)
            
            ttk.Label(day_frame, text=f"{day}:", width=15).pack(side=tk.LEFT)
            
            time_frame = ttk.Frame(day_frame)
            time_frame.pack(side=tk.LEFT)
            
            start_var = tk.StringVar()
            end_var = tk.StringVar()
            
            # Set current values if available
            if day in self.current_workplace.hours_of_operation:
                start, end = self.current_workplace.hours_of_operation[day]
                start_var.set(start)
                end_var.set(end)
            
            ttk.Label(time_frame, text="Start:").pack(side=tk.LEFT, padx=5)
            start_entry = ttk.Entry(time_frame, textvariable=start_var, width=10)
            start_entry.pack(side=tk.LEFT, padx=5)
            
            ttk.Label(time_frame, text="End:").pack(side=tk.LEFT, padx=5)
            end_entry = ttk.Entry(time_frame, textvariable=end_var, width=10)
            end_entry.pack(side=tk.LEFT, padx=5)
            
            hours_entries[day] = (start_var, end_var)
            
            # Add closed checkbox
            closed_var = tk.BooleanVar()
            closed_check = ttk.Checkbutton(day_frame, text="Closed", variable=closed_var)
            closed_check.pack(side=tk.LEFT, padx=10)
            
            # If closed, disable time entries
            def toggle_closed(day=day, start_entry=start_entry, end_entry=end_entry, var=closed_var):
                if var.get():
                    start_entry.config(state="disabled")
                    end_entry.config(state="disabled")
                else:
                    start_entry.config(state="normal")
                    end_entry.config(state="normal")
            
            closed_check.config(command=toggle_closed)
        
        def save_hours():
            hours = {}
            for day, (start_var, end_var) in hours_entries.items():
                start = start_var.get().strip()
                end = end_var.get().strip()
                
                # Validate time format (simple validation)
                try:
                    if start and end:
                        # Simple validation - could be enhanced
                        if ":" not in start or ":" not in end:
                            raise ValueError(f"Invalid time format for {day}")
                        
                        hours[day] = (start, end)
                except ValueError as e:
                    messagebox.showerror("Error", str(e))
                    return
            
            self.current_workplace.hours_of_operation = hours
            messagebox.showinfo("Success", "Hours of operation saved successfully!")
        
        save_hours_button = ttk.Button(hours_frame, text="Save Hours", command=save_hours)
        save_hours_button.pack(pady=20)
        
        # Schedule tab
        schedule_frame = ttk.Frame(notebook)
        notebook.add(schedule_frame, text="Generate Schedule")
        
        ttk.Label(schedule_frame, text="Generate AI Schedule", 
                 font=("Arial", 12, "bold")).pack(pady=10)
        
        param_frame = ttk.LabelFrame(schedule_frame, text="Schedule Parameters")
        param_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Date range
        date_frame = ttk.Frame(param_frame)
        date_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(date_frame, text="Date Range:").pack(side=tk.LEFT, padx=5)
        
        self.start_date_var = tk.StringVar()
        today = datetime.now()
        self.start_date_var.set(today.strftime("%Y-%m-%d"))
        
        ttk.Label(date_frame, text="Start:").pack(side=tk.LEFT, padx=5)
        start_date_entry = ttk.Entry(date_frame, textvariable=self.start_date_var, width=12)
        start_date_entry.pack(side=tk.LEFT, padx=5)
        
        self.end_date_var = tk.StringVar()
        next_week = today + timedelta(days=7)
        self.end_date_var.set(next_week.strftime("%Y-%m-%d"))
        
        ttk.Label(date_frame, text="End:").pack(side=tk.LEFT, padx=5)
        end_date_entry = ttk.Entry(date_frame, textvariable=self.end_date_var, width=12)
        end_date_entry.pack(side=tk.LEFT, padx=5)
        
        # Shift length
        shift_frame = ttk.Frame(param_frame)
        shift_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(shift_frame, text="Default Shift Length (hours):").pack(side=tk.LEFT, padx=5)
        
        self.shift_length_var = tk.StringVar(value="8")
        shift_length_entry = ttk.Entry(shift_frame, textvariable=self.shift_length_var, width=5)
        shift_length_entry.pack(side=tk.LEFT, padx=5)
        
        # Min staff per shift
        staff_frame = ttk.Frame(param_frame)
        staff_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(staff_frame, text="Minimum Staff per Shift:").pack(side=tk.LEFT, padx=5)
        
        self.min_staff_var = tk.StringVar(value="2")
        min_staff_entry = ttk.Entry(staff_frame, textvariable=self.min_staff_var, width=5)
        min_staff_entry.pack(side=tk.LEFT, padx=5)
        
        # Generate button
        def generate_schedule():
            if not self.current_workplace.workers:
                messagebox.showerror("Error", "Please import worker data first!")
                return
            
            try:
                # Get parameters
                start_date = datetime.strptime(self.start_date_var.get(), "%Y-%m-%d")
                end_date = datetime.strptime(self.end_date_var.get(), "%Y-%m-%d")
                shift_length = float(self.shift_length_var.get())
                min_staff = int(self.min_staff_var.get())
                
                # Generate schedule (simplified algorithm)
                schedule = self.generate_ai_schedule(start_date, end_date, shift_length, min_staff)
                
                # Display schedule
                self.schedule_text.delete(1.0, tk.END)
                self.schedule_text.insert(tk.END, "Generated Schedule:\n\n")
                
                for date, shifts in schedule.items():
                    self.schedule_text.insert(tk.END, f"{date.strftime('%Y-%m-%d (%A)')}\n")
                    self.schedule_text.insert(tk.END, "-" * 40 + "\n")
                    
                    for shift, workers in shifts.items():
                        self.schedule_text.insert(tk.END, f"{shift}: {', '.join(workers)}\n")
                    
                    self.schedule_text.insert(tk.END, "\n")
                
                # Save to workplace
                self.current_workplace.shifts = schedule
                messagebox.showinfo("Success", "Schedule generated successfully!")
                
            except Exception as e:
                messagebox.showerror("Error", f"Error generating schedule: {str(e)}")
        
        generate_button = ttk.Button(schedule_frame, text="Generate Schedule", 
                                    command=generate_schedule)
        generate_button.pack(pady=10)
        
        # Schedule display
        schedule_display = ttk.LabelFrame(schedule_frame, text="Generated Schedule")
        schedule_display.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Add scrollbar for schedule
        schedule_scroll = ttk.Scrollbar(schedule_display)
        schedule_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.schedule_text = tk.Text(schedule_display, wrap=tk.WORD,
                                    yscrollcommand=schedule_scroll.set)
        self.schedule_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        schedule_scroll.config(command=self.schedule_text.yview)
        
        # Display existing schedule if available
        if hasattr(self.current_workplace, 'shifts') and self.current_workplace.shifts:
            self.schedule_text.insert(tk.END, "Existing Schedule:\n\n")
            for date, shifts in self.current_workplace.shifts.items():
                if isinstance(date, str):
                    date_str = date
                else:
                    date_str = date.strftime('%Y-%m-%d (%A)')
                    
                self.schedule_text.insert(tk.END, f"{date_str}\n")
                self.schedule_text.insert(tk.END, "-" * 40 + "\n")
                
                for shift, workers in shifts.items():
                    self.schedule_text.insert(tk.END, f"{shift}: {', '.join(workers)}\n")
                
                self.schedule_text.insert(tk.END, "\n")
    
    def generate_ai_schedule(self, start_date, end_date, shift_length, min_staff):
        """
        Generate a schedule based on workplace data and constraints.
        This is a simplified algorithm and could be enhanced with more AI techniques.
        """
        schedule = {}
        current_date = start_date
        
        # Get all workers
        workers = [w["Name"] for w in self.current_workplace.workers]
        
        # For each day in the date range
        while current_date <= end_date:
            day_name = current_date.strftime("%A")
            
            # Skip days when workplace is closed
            if day_name not in self.current_workplace.hours_of_operation:
                current_date += timedelta(days=1)
                continue
            
            # Get operating hours
            start_time_str, end_time_str = self.current_workplace.hours_of_operation.get(
                day_name, ("9:00", "17:00"))
            
            # Parse time strings to datetime objects for calculation
            start_h, start_m = map(int, start_time_str.split(":"))
            end_h, end_m = map(int, end_time_str.split(":"))
            
            # Calculate total operating hours
            total_hours = (end_h - start_h) + (end_m - start_m) / 60
            
            # Calculate number of shifts
            num_shifts = max(1, int(total_hours / shift_length))
            
            # Hours per shift
            hours_per_shift = total_hours / num_shifts
            
            # Generate shifts
            day_shifts = {}
            for i in range(num_shifts):
                shift_start_h = start_h + int(i * hours_per_shift)
                shift_start_m = int((i * hours_per_shift - int(i * hours_per_shift)) * 60)
                
                shift_end_h = start_h + int((i + 1) * hours_per_shift)
                shift_end_m = int(((i + 1) * hours_per_shift - int((i + 1) * hours_per_shift)) * 60)
                
                # Format shift times
                shift_start = f"{shift_start_h:02d}:{shift_start_m:02d}"
                shift_end = f"{shift_end_h:02d}:{shift_end_m:02d}"
                shift_name = f"{shift_start} - {shift_end}"
                
                # Assign workers (simplified - random assignment)
                # In a real AI scheduler, this would consider availability, skills, fairness, etc.
                np.random.seed(int(current_date.timestamp()) + i)  # Deterministic for demo
                assigned_workers = np.random.choice(
                    workers, 
                    size=min(len(workers), min_staff + np.random.randint(0, 3)),
                    replace=False
                ).tolist()
                
                day_shifts[shift_name] = assigned_workers
            
            schedule[current_date] = day_shifts
            current_date += timedelta(days=1)
        
        return schedule
    
    def save_workplaces(self):
        try:
            with open("workplaces.pkl", "wb") as f:
                pickle.dump(self.workplaces, f)
            messagebox.showinfo("Success", "All workplaces saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Error saving workplaces: {str(e)}")
    
    def load_workplaces(self):
        try:
            if os.path.exists("workplaces.pkl"):
                with open("workplaces.pkl", "rb") as f:
                    self.workplaces = pickle.load(f)
        except Exception as e:
            messagebox.showerror("Error", f"Error loading workplaces: {str(e)}")
            self.workplaces = []

def main():
    root = tk.Tk()
    app = WorkplaceSchedulerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
