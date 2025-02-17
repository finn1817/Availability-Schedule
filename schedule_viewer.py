import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import json
import random
import time
from datetime import datetime, timedelta

DATA_FILE = "data.json"

def load_workplaces():
    """Load workplace data from data.json."""
    if not DATA_FILE:
        return []
    try:
        with open(DATA_FILE, "r") as file:
            data = json.load(file)
        return data.get("workplaces", [])
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load workplace data: {e}")
        return []

class WeeklyScheduleGenerator:
    def __init__(self, root):
        """Initialize main GUI components."""
        self.root = root
        self.root.title("Weekly Schedule Generator")
        self.root.geometry("500x400")

        self.workplaces = load_workplaces()
        self.selected_workplace = tk.StringVar()
        self.schedule_data = None
        self.worker_data = None

        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="Select Workplace:", font=('Helvetica', 12)).pack()
        self.workplace_dropdown = ttk.Combobox(main_frame, textvariable=self.selected_workplace, values=[wp['name'] for wp in self.workplaces])
        self.workplace_dropdown.pack(pady=5)

        ttk.Button(main_frame, text="Load Worker Availability", command=self.load_worker_availability).pack(pady=5, fill=tk.X)
        ttk.Button(main_frame, text="Generate Schedule", command=self.generate_schedule).pack(pady=10, fill=tk.X)
    
    def load_worker_availability(self):
        """Load worker availability from an Excel file."""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        try:
            self.worker_data = pd.read_excel(file_path)
            messagebox.showinfo("Success", "Worker availability loaded successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load worker data: {e}")
            self.worker_data = None

    def get_shifts_from_hours(self, hours):
        """Generate shift slots based on opening and closing hours."""
        shifts = []
        try:
            start_time, end_time = hours.lower().replace(" ", "").split("-")
            start_dt = datetime.strptime(start_time, "%I%p")
            end_dt = datetime.strptime(end_time, "%I%p")
            
            if end_dt < start_dt:
                end_dt += timedelta(days=1)  # handle any overnight shifts
            
            current_shift = start_dt
            while current_shift + timedelta(hours=3) <= end_dt:
                shift_end = current_shift + timedelta(hours=3)
                shifts.append(f"{current_shift.strftime('%I %p')} - {shift_end.strftime('%I %p')}")
                current_shift = shift_end
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse shifts: {e}")
        return shifts

    def generate_schedule(self):
        """Generate a weekly schedule dynamically based on workplace hours and worker availability."""
        if self.worker_data is None:
            messagebox.showerror("Error", "Please load worker availability first.")
            return

        workplace_name = self.selected_workplace.get()
        workplace = next((wp for wp in self.workplaces if wp['name'] == workplace_name), None)
        
        if not workplace:
            messagebox.showerror("Error", "Please select a valid workplace.")
            return
        
        shifts_by_day = {}
        schedule = {}
        
        for day, hours in workplace['hours'].items():
            shifts_by_day[day] = self.get_shifts_from_hours(hours)
            schedule[day] = {}
            
            for shift in shifts_by_day[day]:
                available_workers = self.worker_data[(self.worker_data[day].notna()) & (self.worker_data[day].str.lower() != 'na')]
                
                if not available_workers.empty:
                    assigned_worker = available_workers.sample(n=1).iloc[0]
                    schedule[day][shift] = f"{assigned_worker['First Name']} {assigned_worker['Last Name']}"
                else:
                    schedule[day][shift] = "Unassigned"
        
        self.schedule_data = schedule
        messagebox.showinfo("Success", "Schedule generated dynamically based on workplace hours and worker availability!")

if __name__ == "__main__":
    root = tk.Tk()
    app = WeeklyScheduleGenerator(root)
    root.mainloop()
