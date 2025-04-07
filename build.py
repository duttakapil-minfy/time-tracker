#!/usr/bin/env python3
import os
import platform
import subprocess
import sys

def build_executable():
    """Build the executable using PyInstaller"""
    print("Building TimeTracker application...")
    
    # Run PyInstaller with the spec file
    cmd = ["pyinstaller", "timetracker.spec", "--clean"]
    
    try:
        subprocess.check_call(cmd)
        
        # Check if build was successful
        if platform.system() == "Windows":
            exe_path = os.path.join("dist", "TimeTracker.exe")
        elif platform.system() == "Darwin":  # macOS
            exe_path = os.path.join("dist", "TimeTracker.app")
        else:  # Linux
            exe_path = os.path.join("dist", "TimeTracker")
            
        if os.path.exists(exe_path):
            print(f"\nBuild successful! Your application is at: {exe_path}")
            print("\nYou can distribute this executable to others.")
            if platform.system() == "Darwin":
                print("On macOS, you may need to right-click and select 'Open' the first time you run it.")
        else:
            print("\nBuild completed but executable not found at expected location.")
    except subprocess.CalledProcessError as e:
        print(f"Error building executable: {e}")
        return False
        
    return True
    
if __name__ == "__main__":
    # Make sure PyInstaller is installed
    try:
        import PyInstaller
    except ImportError:
        print("PyInstaller not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    build_executable() 