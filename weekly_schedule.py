import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageDraw, ImageFont
from docx import Document

class WeeklyScheduleGenerator:
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
        
        self.schedule_data = None  # will hold the processed schedule data
        self.image_path = "weekly_schedule.png"  # path to save the generated image

    def load_schedule(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        
        try:
            # load the data into a pandas dataframe
            self.schedule_df = pd.read_excel(file_path)

            # make sure excel columns are their
            required_columns = {'First Name', 'Last Name', 'Email', 'Days Available'}
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

        # defining days with their working hours (lounge hours)
        time_slots = {
            "Saturday": "12 PM - 12 AM",
            "Sunday": "12 PM - 12 AM",
            "Monday": "2 PM - 12 AM",
            "Tuesday": "2 PM - 12 AM",
            "Wednesday": "2 PM - 12 AM",
            "Thursday": "2 PM - 12 AM",
            "Friday": "2 PM - 12 AM"
        }

        # prep the schedule dictionary
        weekly_schedule = {day: [] for day in time_slots.keys()}

        # populate the weekly schedule based on "Days Available" tab in excel
        for _, row in self.schedule_df.iterrows():
            for day in time_slots.keys():
                if day in row["Days Available"]:
                    weekly_schedule[day].append(f"{row['First Name']} {row['Last Name']}")

        self.schedule_data = weekly_schedule

        # make the calendar style pic
        self.create_calendar_image(weekly_schedule, time_slots)
        messagebox.showinfo("Success", "Weekly schedule generated! Image saved as 'weekly_schedule.png'.")

    def create_calendar_image(self, weekly_schedule, time_slots):
        # pic settings
        width, height = 1000, 700
        header_height = 100
        color_bg = (255, 255, 255)
        color_header = (70, 130, 180)
        color_text = (0, 0, 0)

        # make pic
        img = Image.new("RGB", (width, height), color_bg)
        draw = ImageDraw.Draw(img)

        # add Title
        font_title = ImageFont.truetype("arial.ttf", 40)
        draw.text((width // 2 - 150, 10), "Weekly Schedule", fill=color_header, font=font_title)

        # add the days and schedules
        font_body = ImageFont.truetype("arial.ttf", 20)
        row_height = (height - header_height) // 7

        for i, (day, workers) in enumerate(weekly_schedule.items()):
            y_start = header_height + i * row_height

            # draw day header
            draw.rectangle([0, y_start, width, y_start + row_height], outline=color_header, width=3)
            day_text = f"{day} ({time_slots[day]})"
            draw.text((10, y_start + 10), day_text, fill=color_header, font=font_body)

            # add worker names
            worker_text = ", ".join(workers) if workers else "No workers available"
            draw.text((20, y_start + 40), worker_text, fill=color_text, font=font_body)

        # save the image
        img.save(self.image_path)

    def save_word_file(self):
        if not self.schedule_data:
            messagebox.showerror("Error", "Please generate the schedule first.")
            return

        try:
            # make word document
            doc = Document()
            doc.add_heading("Weekly Schedule", level=1)

            for day, workers in self.schedule_data.items():
                doc.add_heading(day, level=2)
                doc.add_paragraph(", ".join(workers) if workers else "No workers available")
            
            # save Word file
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
