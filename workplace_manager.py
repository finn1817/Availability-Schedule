import tkinter as tk
from tkinter import messagebox, simpledialog, filedialog
import pandas as pd
import json
from weekly_schedule import WeeklyScheduleGenerator
from schedule_viewer import ScheduleViewer


class WorkplaceManager:
    def __init__(self, root, data, save_callback):
        self.root = tk.Toplevel(root)
        self.root.title("Manage Workplaces & Schedules")
        self.root.geometry("800x600")
        self.data = data
        self.save_callback = save_callback

        # main layout
        self.setup_layout()

    def setup_layout(self):
        """Create the main layout of the Workplace Manager."""
        # workplace List Section
        workplace_frame = tk.Frame(self.root)
        workplace_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)

        tk.Label(workplace_frame, text="Workplaces", font=("Arial", 14, "bold")).pack()

        self.workplace_listbox = tk.Listbox(workplace_frame)
        self.workplace_listbox.pack(fill=tk.BOTH, expand=True)
        self.workplace_listbox.bind("<<ListboxSelect>>", self.display_workplace_details)

        action_frame = tk.Frame(self.root)
        action_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=10, pady=10)

        tk.Button(action_frame, text="Add Workplace", command=self.add_workplace).pack(fill=tk.X, pady=5)
        tk.Button(action_frame, text="Delete Workplace", command=self.delete_workplace).pack(fill=tk.X, pady=5)
        tk.Button(action_frame, text="Modify Hours", command=self.modify_workplace_hours).pack(fill=tk.X, pady=5)

        # schedule Options
        schedule_label = tk.Label(action_frame, text="Schedule Tools", font=("Arial", 14, "bold"))
        schedule_label.pack(pady=(20, 10))
        
        tk.Button(action_frame, text="View Schedule", command=self.view_schedule).pack(fill=tk.X, pady=5)
        tk.Button(action_frame, text="Generate Weekly Schedule", command=self.generate_weekly_schedule).pack(fill=tk.X, pady=5)

        tk.Button(action_frame, text="Save All Changes", command=self.save_data).pack(fill=tk.X, pady=10)

        self.refresh_workplace_list()

    def refresh_workplace_list(self):
        """Update the workplace list."""
        self.workplace_listbox.delete(0, tk.END)
        for workplace in self.data["workplaces"]:
            self.workplace_listbox.insert(tk.END, workplace["name"])

    def add_workplace(self):
        """Add a new workplace to the list."""
        new_name = simpledialog.askstring("Add Workplace", "Enter workplace name:")
        if new_name:
            self.data["workplaces"].append({"name": new_name, "hours": {}})
            self.refresh_workplace_list()

    def delete_workplace(self):
        """Remove a selected workplace."""
        selected_index = self.workplace_listbox.curselection()
        if selected_index:
            del self.data["workplaces"][selected_index[0]]
            self.refresh_workplace_list()
            self.save_data()

    def modify_workplace_hours(self):
        """Edit working hours for a specific workplace."""
        selected_index = self.workplace_listbox.curselection()
        if not selected_index:
            messagebox.showerror("Error", "Please select a workplace to edit.")
            return

        workplace = self.data["workplaces"][selected_index[0]]

        for day in ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]:
            default_hours = workplace["hours"].get(day, "9 AM - 5 PM")
            new_hours = simpledialog.askstring(f"Modify Hours ({day})", f"Enter working hours for {day}:", initialvalue=default_hours)
            if new_hours:
                workplace["hours"][day] = new_hours

        messagebox.showinfo("Success", f"Updated hours for workplace '{workplace['name']}'.")

    def display_workplace_details(self, event):
        """Display details of a selected workplace."""
        selected_index = self.workplace_listbox.curselection()
        if selected_index:
            workplace = self.data["workplaces"][selected_index[0]]
            messagebox.showinfo("Workplace Details", json.dumps(workplace, indent=4))

    def view_schedule(self):
        """Launch the Schedule Viewer."""
        ScheduleViewer(self.root)

    def generate_weekly_schedule(self):
        """Launch the Weekly Schedule Generator."""
        WeeklyScheduleGenerator(self.root)

    def save_data(self):
        """Save changes to data."""
        self.save_callback()
        self.refresh_workplace_list()
        messagebox.showinfo("Save", "Workplaces and schedules saved successfully!")
