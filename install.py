import os
import subprocess
import sys

def check_python_installation():
    try:
        python_version = subprocess.check_output(["python", "--version"]).decode("utf-8").strip()
        print(f"Python version detected: {python_version}")
        return True
    except FileNotFoundError:
        print("Python is not installed. Please install it and try again.")
        return False

def install_packages():
    if not check_python_installation():
        return

    print("Installing required Python packages...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "pip"])  # upgrades pip too
        required_packages = ["pandas", "openpyxl", "pillow", "python-docx"]
        subprocess.check_call([sys.executable, "-m", "pip", "install"] + required_packages)
        print("Installation complete!")
    except Exception as e:
        print(f"Error during installation: {e}")

if __name__ == "__main__":
    install_packages()
