# ZKTeco Attendance Report Tool

A Python-based tool to extract attendance data from ZKTeco biometric machines and generate monthly Excel reports with minimal user interaction.

This project was created to help non-technical users generate attendance reports with a single click, without navigating complex company software.

## Features
- Connects to ZKTeco biometric device
- Handles missing clock-in / clock-out automatically
- Calculates daily and monthly worked hours
- Exports clean, formatted Excel reports
- Supports Nepali date and time

## For Normal Users (Recommended)
No Python knowledge required.

1. Download the EXE folder
2. Open the folder
3. Double-click `zkteco.exe`
4. Enter the month when prompted
5. Excel report is generated automatically

(This EXE is built using PyInstaller `--onedir`)

## For Developers
If you want to run or modify the source code:

```bash
pip install -r requirements.txt
python zkteco.py


Why this project?

Manual attendance calculation using ZKTeco software is time-consuming and error-prone.
This tool simplifies the workflow and is designed for real-world HR usage.

Author
Abhishek Rai