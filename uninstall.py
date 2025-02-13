import os

def uninstall_packages():
    print("Removing installed Python packages...")
    os.system("pip uninstall -y pandas openpyxl pillow python-docx")
    print("Uninstall complete!")

if __name__ == "__main__":
    uninstall_packages()
