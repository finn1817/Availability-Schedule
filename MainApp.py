import os
import subprocess
import tkinter as tk
from tkinter import messagebox, ttk
from tkinter.font import Font

class ModernButton(ttk.Button):
    """Custom styled button class"""
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.configure(style='Modern.TButton')

class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Schedule Manager")
        self.root.geometry("500x400")
        self.root.configure(bg='#f0f0f0')  # for light gray background
        
        # make window minimum size
        self.root.minsize(400, 300)
        
        # set weights for responsive layout
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)
        
        # make and configure styles
        self.setup_styles()
        
        # make the main container frame
        self.main_frame = ttk.Frame(root, style='Main.TFrame')
        self.main_frame.grid(row=0, column=0, sticky='nsew', padx=21, pady=21)
        self.main_frame.grid_columnconfigure(0, weight=1)
        
        # make and pack widgets
        self.create_widgets()
        
    def setup_styles(self):
        """Configure custom styles for the application"""
        # set up modern button style
        style = ttk.Style()
        style.configure(
            'Modern.TButton',
            padding=10,
            font=('Helvetica', 9),
            background='#2196F3',  # material design blue
            relief='flat'
        )
        
        # configure frame style
        style.configure(
            'Main.TFrame',
            background='#ffffff'  # white background
        )
        
        # configure title style
        style.configure(
            'Title.TLabel',
            font=('Helvetica', 24, 'bold'),
            background='#ffffff',
            foreground='#333333'  # dark gray text
        )
        
    def create_widgets(self):
        """Create and arrange all widgets"""
        # title with modern font and spacing
        title_label = ttk.Label(
            self.main_frame,
            text="Schedule Manager",
            style='Title.TLabel'
        )
        title_label.grid(row=0, column=0, pady=(20, 40), sticky='n')
        
        # button container for consistent layout
        button_frame = ttk.Frame(self.main_frame, style='Main.TFrame')
        button_frame.grid(row=1, column=0, sticky='n')
        button_frame.grid_columnconfigure(0, weight=1)
        
        # schedule viewer button
        self.schedule_viewer_button = ModernButton(
            button_frame,
            text="📅 View Schedule",
            command=self.open_schedule_viewer,
            width=25
        )
        self.schedule_viewer_button.grid(row=0, column=0, pady=10, padx=20)
        
        # weekly schedule maker button
        self.weekly_schedule_button = ModernButton(
            button_frame,
            text="✏️ Create Weekly Schedule",
            command=self.open_weekly_schedule,
            width=25
        )
        self.weekly_schedule_button.grid(row=1, column=0, pady=10, padx=20)
        
        # exit button with different styling
        self.exit_button = ModernButton(
            button_frame,
            text="❌ Exit",
            command=self.exit_app,
            width=25
        )
        self.exit_button.grid(row=2, column=0, pady=(30, 10), padx=20)
        
    def open_schedule_viewer(self):
        """Opens the Schedule Viewer script in a new process."""
        try:
            if os.path.exists("schedule_viewer.py"):
                subprocess.Popen(["python", "schedule_viewer.py"])
            else:
                messagebox.showerror(
                    "File Not Found",
                    "The schedule_viewer.py file is missing."
                )
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Schedule Viewer: {e}")
            
    def open_weekly_schedule(self):
        """Opens the Weekly Schedule Maker script in a new process."""
        try:
            if os.path.exists("weekly_schedule.py"):
                subprocess.Popen(["python", "weekly_schedule.py"])
            else:
                messagebox.showerror(
                    "File Not Found",
                    "The weekly_schedule.py file is missing."
                )
        except Exception as e:
            messagebox.showerror(
                "Error",
                f"Failed to open Weekly Schedule Maker: {e}"
            )
            
    def exit_app(self):
        """Exits the main application."""
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = MainApp(root)
    root.mainloop()
