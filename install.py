import os

print("Installing required Python packages...")
os.system("pip install pandas openpyxl")
print("Installation complete!")

if __name__ == "__main__":
    install_packages()
