import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageDraw, ImageFont
from docx import Document
import random
import time

class WeeklyScheduleGenerator:
    def __init__(self, root):
        """Initialize main GUI components."""
        self.root = root
        self.root.title("Weekly Schedule Generator")
        self.root.geometry("500x400")
        
        # config styles
        style = ttk.Style()
        style.configure('TButton', padding=10)
        
        # making the main frame
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # title label
        title_label = ttk.Label(main_frame, text="Weekly Schedule Generator", 
                              font=('Helvetica', 16, 'bold'))
        title_label.pack(pady=(0, 20))

        # buttons for the app
        self.load_button = ttk.Button(main_frame, text="Load Excel Schedule", 
                                    command=self.load_schedule)
        self.load_button.pack(pady=10, fill=tk.X)

        self.generate_availability_button = ttk.Button(main_frame, 
                                                     text="Generate Availability", 
                                                     command=self.generate_availability)
        self.generate_availability_button.pack(pady=10, fill=tk.X)

        self.generate_schedule_button = ttk.Button(main_frame, 
                                                 text="Generate Schedule", 
                                                 command=self.generate_schedule)
        self.generate_schedule_button.pack(pady=10, fill=tk.X)

        self.save_button = ttk.Button(main_frame, text="Save Word File", 
                                    command=self.save_word_file)
        self.save_button.pack(pady=10, fill=tk.X)

        # data containers
        self.schedule_df = None
        self.availability_data = None
        self.schedule_data = None
        self.image_path = "final_schedule.png"

    def load_schedule(self):
        """Loads the Excel schedule file into a DataFrame."""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return

        try:
            # load data into a pandas DataFrame
            self.schedule_df = pd.read_excel(file_path)

            # making sure all columns are there
            required_columns = {'First Name', 'Last Name', 'Email', 'Sunday', 'Monday', 
                              'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'}
            if not required_columns.issubset(self.schedule_df.columns):
                messagebox.showerror("Error", 
                                   f"Excel file must contain the following columns: {', '.join(required_columns)}")
                return

            messagebox.showinfo("Success", "Schedule file loaded successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load schedule file: {e}")
            self.schedule_df = None

    def generate_availability(self):
        """Processes the loaded schedule to create availability data per day and shift."""
        if self.schedule_df is None:
            messagebox.showerror("Error", "Please load a schedule file first.")
            return

        # defining available shifts and times
        shifts = {
            "Sunday": ["12 PM - 4 PM", "4 PM - 7 PM", "7 PM - 10 PM", "10 PM - 12 AM"],
            "Monday": ["2 PM - 5 PM", "5 PM - 8 PM", "8 PM - 12 AM"],
            "Tuesday": ["2 PM - 5 PM", "5 PM - 8 PM", "8 PM - 12 AM"],
            "Wednesday": ["2 PM - 6 PM", "6 PM - 9 PM", "9 PM - 12 AM"],
            "Thursday": ["2 PM - 4 PM", "4 PM - 8 PM", "8 PM - 12 AM"],
            "Friday": ["2 PM - 7 PM", "7 PM - 9 PM", "9 PM - 12 AM"],
            "Saturday": ["12 PM - 4 PM", "4 PM - 8 PM", "8 PM - 12 AM"]
        }

        # making a nested dictionary for availability
        availability = {day: {shift: [] for shift in shifts[day]} for day in shifts.keys()}

        for _, row in self.schedule_df.iterrows():
            for day in availability.keys():
                if pd.notna(row[day]) and row[day].lower() != 'na':
                    for shift in availability[day]:
                        availability[day][shift].append(f"{row['First Name']} {row['Last Name']}")

        self.availability_data = availability
        messagebox.showinfo("Success", "Availability data created successfully!")

    def generate_schedule(self):
        """Generates a balanced schedule based on availability data."""
        if not self.availability_data:
            messagebox.showerror("Error", "Please generate availability first.")
            return

        # start worker hours tracker
        worker_hours = {}  # key: worker name, value: total hours assigned
        
        # helper to calculate shift duration in hours
        def calculate_shift_duration(shift_time):
            """Calculate duration of a shift in hours."""
            start, end = shift_time.split(" - ")

            def convert_to_24hr(time_str):
                time = int(time_str.split()[0])
                if "PM" in time_str and time != 12:
                    time += 12
                elif "AM" in time_str and time == 12:
                    time = 0
                return time

            start_hour = convert_to_24hr(start)
            end_hour = convert_to_24hr(end)
            if end_hour < start_hour:  # account for overnight shift
                end_hour += 24
            return end_hour - start_hour

        # start schedule dictionary
        schedule = {day: {shift: None for shift in shifts} 
                    for day, shifts in self.availability_data.items()}

        # assign workers while trying to balance hours
        for day, shifts in self.availability_data.items():
            for shift, workers in shifts.items():
                if not workers:  # skip if no workers are available for this shift
                    continue

                # get the duration of the shift
                shift_duration = calculate_shift_duration(shift)

                # start workers not in worker_hours with 0 total hours
                for worker in workers:
                    if worker not in worker_hours:
                        worker_hours[worker] = 0

                # sort out workers by their total assigned hours (ascending)
                workers_sorted = sorted(workers, key=lambda w: worker_hours[w])

                # assign the available worker with the least total hours
                assigned_worker = None
                for worker in workers_sorted:
                    assigned_worker = worker
                    break

                # give the worker to the shift
                schedule[day][shift] = assigned_worker
                
                if assigned_worker:
                    # update workers total hours
                    worker_hours[assigned_worker] += shift_duration

        self.schedule_data = schedule
        self.create_calendar_image(schedule)
        messagebox.showinfo("Success", "Schedule generated successfully!")

    def create_calendar_image(self, data, title="Weekly Schedule"):
        """Generates a PNG visual representation of the schedule."""
        # pic settings
        width, height = 1200, 1000
        header_height = 100
        color_bg = (255, 255, 255)
        color_header = (70, 130, 180)
        color_text = (0, 0, 0)

        # make the image
        img = Image.new("RGB", (width, height), color_bg)
        draw = ImageDraw.Draw(img)

        try:
            font_title = ImageFont.truetype("arial.ttf", 40)
            font_body = ImageFont.truetype("arial.ttf", 20)
        except:
            # fallback to default font if arial isn't available
            font_title = ImageFont.load_default()
            font_body = ImageFont.load_default()

        # add title
        draw.text((width // 2 - 200, 10), title, fill=color_header, font=font_title)

        # draw schedule data
        row_height = (height - header_height) // len(data)

        for i, (day, shift_data) in enumerate(data.items()):
            y_start = header_height + i * row_height

            # draw day header
            draw.text((10, y_start), day, fill=color_header, font=font_body)

            # add shifts and workers
            y_shift = y_start + 40
            for shift, worker in shift_data.items():
                shift_text = f"{shift}: {worker if worker else 'Unassigned'}"
                draw.text((40, y_shift), shift_text, fill=color_text, font=font_body)
                y_shift += 30

        img.save(self.image_path)

    def save_word_file(self):
        """Saves the generated schedule and worker summary to a Word document."""
        if not self.schedule_data:
            messagebox.showerror("Error", "Please generate the final schedule first.")
            return

        try:
            doc = Document()
            
            # add title with formatting
            title = doc.add_heading("Weekly Schedule", level=1)
            title.alignment = 1  # center alignment
            doc.add_paragraph()
            
            # track worker hours
            worker_hours = {}
            
            def calculate_shift_duration(shift_time):
                """Calculate duration of a shift in hours."""
                start, end = shift_time.split(" - ")
                
                def convert_to_24hr(time_str):
                    time = int(time_str.split()[0])
                    if "PM" in time_str and time != 12:
                        time += 12
                    elif "AM" in time_str and time == 12:
                        time = 0
                    return time
                
                start_hour = convert_to_24hr(start)
                end_hour = convert_to_24hr(end)
                
                if end_hour < start_hour:
                    end_hour += 24
                    
                return end_hour - start_hour

            # add each day's schedule
            for day, shifts in self.schedule_data.items():
                # add day header
                day_heading = doc.add_heading(day, level=2)
                
                # make table for shifts
                table = doc.add_table(rows=1, cols=2)
                table.style = 'Table Grid'
                
                # set headers
                header_cells = table.rows[0].cells
                header_cells[0].text = "Shift Time"
                header_cells[1].text = "Worker"
                
                # add shifts to table
                for shift, worker in shifts.items():
                    row_cells = table.add_row().cells
                    row_cells[0].text = shift
                    row_cells[1].text = worker if worker else 'Unassigned'
                    
                    # calculate hours
                    if worker:
                        duration = calculate_shift_duration(shift)
                        worker_hours[worker] = worker_hours.get(worker, 0) + duration
                
                doc.add_paragraph()  # add spacing

            # add hours summary
            doc.add_heading("Weekly Hours Summary", level=2)
            summary_table = doc.add_table(rows=1, cols=2)
            summary_table.style = 'Table Grid'
            
            # set up summary headers
            header_cells = summary_table.rows[0].cells
            header_cells[0].text = "Worker Name"
            header_cells[1].text = "Total Hours"
            
            # add worker hours (sorted by hours)
            sorted_workers = sorted(worker_hours.items(), key=lambda x: x[1], reverse=True)
            for worker, hours in sorted_workers:
                row_cells = summary_table.add_row().cells
                row_cells[0].text = worker
                row_cells[1].text = f"{hours} hours"
            
            # add generation timestamp
            doc.add_paragraph()
            footer = doc.add_paragraph(f"Generated on: {time.strftime('%Y-%m-%d %H:%M:%S')}")
            footer.alignment = 1

            # save the file
            file_path = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word files", "*.docx")],
                initialfile="Weekly_Schedule.docx"
            )
            if file_path:
                doc.save(file_path)
                messagebox.showinfo("Success", "Schedule saved successfully!")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save Word file: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = WeeklyScheduleGenerator(root)
    root.mainloop()
