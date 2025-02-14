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

def uninstall_packages():
    if not check_python_installation():
        return

    print("Removing installed Python packages...")
    try:
        required_packages = ["pandas", "openpyxl", "pillow", "python-docx"]
        subprocess.check_call([sys.executable, "-m", "pip", "uninstall", "-y"] + required_packages)
        print("Uninstall complete!")
    except Exception as e:
        print(f"Error during uninstallation: {e}")

if __name__ == "__main__":
    uninstall_packages()
