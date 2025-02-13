import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

class ScheduleViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("Schedule Viewer")
        self.root.geometry("600x400")
        
        self.load_button = tk.Button(root, text="Load Schedule", command=self.load_schedule)
        self.load_button.pack(pady=10)
        
        self.day_label = tk.Label(root, text="Select Day:")
        self.day_label.pack()
        
        self.day_dropdown = ttk.Combobox(root, values=["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"])
        self.day_dropdown.pack()
        
        self.text_area = tk.Text(root, height=10, width=70)
        self.text_area.pack()
    
    def load_schedule(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        
        try:
            self.df = pd.read_excel(file_path)
            selected_day = self.day_dropdown.get()
            if not selected_day:
                messagebox.showerror("Error", "Please select a day")
                return
            
            available_workers = self.df[self.df['Days Available'].str.contains(selected_day, na=False)]
            
            self.text_area.delete("1.0", tk.END)
            if not available_workers.empty:
                for _, row in available_workers.iterrows():
                    self.text_area.insert(tk.END, f"{row['First Name']} {row['Last Name']} - Not Available: {row['Time not Available']}\n")
            else:
                self.text_area.insert(tk.END, "No workers available for this day.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel file: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ScheduleViewer(root)
    root.mainloop()
