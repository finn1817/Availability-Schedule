import os

print("Removing installed Python packages...")
os.system("pip uninstall -y pandas openpyxl")
print("Uninstall complete!")
