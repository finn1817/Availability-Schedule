import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

class ScheduleViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("Schedule Viewer")
        self.root.geometry("700x600")
        self.current_file_path = None  # keep the last Excel file path

        self.load_button = tk.Button(root, text="Load Schedule", command=self.load_schedule)
        self.load_button.pack(pady=10)

        self.update_button = tk.Button(root, text="Update", command=self.update_schedule)
        self.update_button.pack(pady=10)

        self.day_label = tk.Label(root, text="Select Day:")
        self.day_label.pack()

        self.day_dropdown = ttk.Combobox(root, values=["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"])
        self.day_dropdown.pack()

        self.not_available_label = tk.Label(root, text="Not Available Workers:")
        self.not_available_label.pack()
        self.not_available_text = tk.Text(root, height=10, width=80)
        self.not_available_text.pack(pady=5)

        self.available_label = tk.Label(root, text="Available Workers:")
        self.available_label.pack()
        self.available_text = tk.Text(root, height=10, width=80)
        self.available_text.pack(pady=5)

    def load_schedule(self):
        self.current_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not self.current_file_path:
            return

        self.update_schedule()

    def update_schedule(self):
        if not self.current_file_path:
            messagebox.showerror("Error", "No schedule file loaded. Please load a file first.")
            return

        try:
            self.df = pd.read_excel(self.current_file_path)
            selected_day = self.day_dropdown.get()

            if not selected_day:
                messagebox.showerror("Error", "Please select a day.")
                return

            required_columns = {'First Name', 'Last Name', 'Email', 'Days Available', 'Time Available on Days Available', 'Time not Available'}
            if not required_columns.issubset(self.df.columns):
                messagebox.showerror("Error", f"Excel file must contain the following columns: {', '.join(required_columns)}")
                return

            self.not_available_text.delete("1.0", tk.END)
            self.available_text.delete("1.0", tk.END)

            selected_day = selected_day.strip()
            available_workers = self.df[
                self.df['Days Available'].str.contains(fr'\b{selected_day}\b', case=False, na=False)
            ]
            not_available_workers = self.df[~self.df['Email'].isin(available_workers['Email'])]

            # show workers who are not available
            if not not_available_workers.empty:
                self.not_available_text.insert(tk.END, f"The following workers are NOT AVAILABLE on {selected_day}:\n")
                self.not_available_text.insert(tk.END, "-" * 50 + "\n")
                for _, row in not_available_workers.iterrows():
                    self.not_available_text.insert(tk.END, f"{row['First Name']} {row['Last Name']} - Not Available\n")
            else:
                self.not_available_text.insert(tk.END, f"No workers marked as 'Not Available' for {selected_day}.\n")

            # show workers who are available
            if not available_workers.empty:
                self.available_text.insert(tk.END, f"The following workers are AVAILABLE on {selected_day}:\n")
                self.available_text.insert(tk.END, "-" * 50 + "\n")
                for _, row in available_workers.iterrows():
                    self.available_text.insert(
                        tk.END,
                        f"{row['First Name']} {row['Last Name']} - Email: {row['Email']} - Time Available: {row['Time Available on Days Available']}\n"
                    )
            else:
                self.available_text.insert(tk.END, f"No workers marked as 'Available' for {selected_day}.\n")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel file or process data: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ScheduleViewer(root)
    root.mainloop()
