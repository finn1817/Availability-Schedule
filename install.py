import os

def install_packages():
    print("Installing required Python packages...")
    os.system("pip install pandas openpyxl pillow python-docx")
    print("Installation complete!")

if __name__ == "__main__":
    install_packages()
