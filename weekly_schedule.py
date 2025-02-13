import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document

class WeeklyScheduleGenerator:
    MAX_HOURS = 4  # max hours allowed per shift
    MAX_SHIFTS_PER_DAY = 1  # limit to one shift per person per day

    def __init__(self, root):
        self.root = root
        self.root.title("Weekly Schedule Generator")
        self.root.geometry("400x200")
        
        self.load_button = tk.Button(root, text="Load Excel Schedule", command=self.load_schedule)
        self.load_button.pack(pady=10)
        
        self.generate_button = tk.Button(root, text="Generate Weekly Schedule", command=self.generate_schedule)
        self.generate_button.pack(pady=10)
        
        self.save_button = tk.Button(root, text="Save Word File", command=self.save_word_file)
        self.save_button.pack(pady=10)
        
        self.schedule_data = None

    def load_schedule(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        
        try:
            self.schedule_df = pd.read_excel(file_path)

            required_columns = {'First Name', 'Last Name', 'Days Available', 'Shift Hours', 'Time available on Days Available'}
            if not required_columns.issubset(self.schedule_df.columns):
                messagebox.showerror("Error", f"Excel file must contain the following columns: {', '.join(required_columns)}")
                return

            messagebox.showinfo("Success", "Schedule file loaded successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load schedule file: {e}")
            self.schedule_df = None

    def generate_schedule(self):
        if self.schedule_df is None:
            messagebox.showerror("Error", "Please load a schedule file first.")
            return
        
        time_slots = {
            "Saturday": ["12 PM - 4 PM", "4 PM - 8 PM", "8 PM - 12 AM"],
            "Sunday": ["12 PM - 4 PM", "4 PM - 8 PM", "8 PM - 12 AM"],
            "Monday": ["2 PM - 6 PM", "6 PM - 9 PM", "9 PM - 12 AM"],
            "Tuesday": ["2 PM - 6 PM", "6 PM - 9 PM", "9 PM - 12 AM"],
            "Wednesday": ["2 PM - 6 PM", "6 PM - 9 PM", "9 PM - 12 AM"],
            "Thursday": ["2 PM - 6 PM", "6 PM - 9 PM", "9 PM - 12 AM"],
            "Friday": ["2 PM - 6 PM", "6 PM - 9 PM", "9 PM - 12 AM"]
            # add in any other times / shifts here
        }

        weekly_schedule = {day: [] for day in time_slots.keys()}

        for _, row in self.schedule_df.iterrows():
            if row['Shift Hours'] > self.MAX_HOURS:
                continue  # skip anyone violating max hours
            
            for day, slots in time_slots.items():
                if day in row['Days Available']:
                    available_times = row['Time available on Days Available'].split(",")  # split available times
                    for slot in slots:
                        if slot.strip() in available_times and len(weekly_schedule[day]) < self.MAX_SHIFTS_PER_DAY:
                            weekly_schedule[day].append(f"{row['First Name']} {row['Last Name']} - {slot.strip()}")
                            break

        self.schedule_data = weekly_schedule
        print(self.schedule_data)  # for debugging

    def save_word_file(self):
        if not self.schedule_data:
            messagebox.showerror("Error", "Please generate the schedule first.")
            return

        try:
            doc = Document()
            doc.add_heading("Weekly Schedule", level=1)

            for day, workers in self.schedule_data.items():
                doc.add_heading(day, level=2)
                doc.add_paragraph(", ".join(workers) if workers else "No workers available")
            
            file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
            if file_path:
                doc.save(file_path)
                messagebox.showinfo("Success", f"Word file saved successfully at {file_path}.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save Word file: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = WeeklyScheduleGenerator(root)
    root.mainloop()
