import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageDraw, ImageFont
from docx import Document

class WeeklyScheduleGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Weekly Schedule Generator")
        self.root.geometry("400x250")
        
        self.load_button = tk.Button(root, text="Load Excel Schedule", command=self.load_schedule)
        self.load_button.pack(pady=10)
        
        self.generate_button = tk.Button(root, text="Generate Weekly Schedule", command=self.generate_schedule)
        self.generate_button.pack(pady=10)
        
        self.save_button = tk.Button(root, text="Save Word File", command=self.save_word_file)
        self.save_button.pack(pady=10)
        
        self.image_button = tk.Button(root, text="Save Image File", command=self.save_image_file)
        self.image_button.pack(pady=10)
        
        self.schedule_data = None  # will hold the processed schedule data
        self.image_path = "weekly_schedule.png"  # path to save the generated picture

    def load_schedule(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        
        try:
            # loading data into a pandas dataframe
            self.schedule_df = pd.read_excel(file_path)

            # main important columns
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

        # make the shifts and their times
        shifts = {
            "Sunday": ["12 PM - 4 PM", "4 PM - 7 PM", "7 PM - 10 PM", "10 PM - 12 AM"],
            "Monday": ["2 PM - 5 PM", "5 PM - 8 PM", "8 PM - 12 AM"],
            "Tuesday": ["2 PM - 5 PM", "5 PM - 8 PM", "8 PM - 12 AM"],
            "Wednesday": ["2 PM - 6 PM", "6 PM - 9 PM", "9 PM - 12 AM"],
            "Thursday": ["2 PM - 4 PM", "4 PM - 8 PM", "8 PM - 12 AM"],
            "Friday": ["2 PM - 7 PM", "7 PM - 9 PM", "9 PM - 12 AM"],
            "Saturday": ["12 PM - 4 PM", "4 PM - 8 PM", "8 PM - 12 AM"]
        }

        # prep schedule dictionary
        weekly_schedule = {day: {shift: [] for shift in shifts[day]} for day in shifts.keys()}

        # populates the schedule based on "Days Available"
        for _, row in self.schedule_df.iterrows():
            for day in weekly_schedule.keys():
                if day in row["Days Available"]:
                    for shift in weekly_schedule[day]:
                        weekly_schedule[day][shift].append(f"{row['First Name']} {row['Last Name']}")

        # check for enough coverage
        for day, shifts in weekly_schedule.items():
            for shift, workers in shifts.items():
                if len(workers) < 1:  # not enough workers for the shift
                    available_workers = [f"{row['First Name']} {row['Last Name']}" for _, row in self.schedule_df.iterrows() if day in row["Days Available"]]
                    # adding the first available worker again (if any)
                    while len(workers) < 1 and available_workers:
                        workers.append(available_workers.pop(0))

        self.schedule_data = weekly_schedule

        # making calendar pic
        self.create_calendar_image(weekly_schedule, shifts)
        messagebox.showinfo("Success", "Weekly schedule generated! Image saved as 'weekly_schedule.png'.")

    def create_calendar_image(self, weekly_schedule, shifts):
        # picture settings
        width, height = 1200, 1000
        header_height = 100
        color_bg = (255, 255, 255)
        color_header = (70, 130, 180)
        color_text = (0, 0, 0)

        # make the picture
        img = Image.new("RGB", (width, height), color_bg)
        draw = ImageDraw.Draw(img)

        # adding title
        font_title = ImageFont.truetype("arial.ttf", 40)
        draw.text((width // 2 - 200, 10), "Weekly Schedule", fill=color_header, font=font_title)

        # Draw the schedule by day and shift
        font_body = ImageFont.truetype("arial.ttf", 20)
        row_height = (height - header_height) // len(weekly_schedule)

        for i, (day, shift_data) in enumerate(weekly_schedule.items()):
            y_start = header_height + i * row_height

            # draw day header
            day_text = day
            draw.text((10, y_start + 10), day_text, fill=color_header, font=font_body)

            # add shifts and workers
            y_shift = y_start + 40
            for shift, workers in shift_data.items():
                shift_text = f"{shift}: {', '.join(workers) if workers else 'No workers available'}"
                draw.text((40, y_shift), shift_text, fill=color_text, font=font_body)
                y_shift += 30

        # Save the image
        img.save(self.image_path)

    def save_word_file(self):
        if not self.schedule_data:
            messagebox.showerror("Error", "Please generate the schedule first.")
            return

        try:
            # make the Word document
            doc = Document()
            doc.add_heading("Weekly Schedule", level=1)

            for day, shift_data in self.schedule_data.items():
                doc.add_heading(day, level=2)
                for shift, workers in shift_data.items():
                    doc.add_paragraph(f"{shift}: {', '.join(workers) if workers else 'No workers available'}")
            
            # save as a Word file
            file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
            if file_path:
                doc.save(file_path)
                messagebox.showinfo("Success", f"Word file saved successfully at {file_path}.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save Word file: {e}")

    def save_image_file(self):
        if self.schedule_data is None:
            messagebox.showerror("Error", "Please generate the schedule first.")
            return
        messagebox.showinfo("Success", f"Image saved as {self.image_path}")

if __name__ == "__main__":
    root = tk.Tk()
    app = WeeklyScheduleGenerator(root)
    root.mainloop()
