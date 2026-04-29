# IPOT Automation

Automates downloading Broker Summary by Stock CSV files from the IPOT (Indonesia Stock Exchange) application for a specified date range.

## Requirements

- Python 3.x
- Windows OS (required for IPOT application automation)
- Virtual environment (already set up in `init/` folder)

## Dependencies

The following Python packages are required:
- pywinauto
- pyautogui
- keyboard

## Setup and Installation

1. **Activate the virtual environment:**
   ```bash
   init\Scripts\activate.bat
   ```

2. **Install dependencies:**
   ```bash
   pip install pywinauto pyautogui keyboard
   ```

3. **Configure the script:**
   - Open `main.py`
   - Edit the following variables at the bottom of the file:
     - `DATE_START`: Start date for data download (format: YYYY-MM-DD)
     - `DATE_END`: End date for data download (format: YYYY-MM-DD)
     - `STOCK_CODE`: Stock code to query (default: "INCO")
     - `SAVE_FOLDER`: Folder where CSV files will be saved (default: "C:\IPOT_Data")

   Make sure the save folder exists before running the script.

## Running the Automation

1. Ensure the IPOT application is installed and running on your system.

2. Run the script:
   ```bash
   python main.py
   ```

3. **Important Notes:**
   - Keep your hands off the mouse and keyboard while the script runs
   - The script will automatically skip weekends (no trading data available)
   - If a modal dialog appears unexpectedly, the script will pause - check the screen
   - Press Ctrl+C in the terminal to abort the script if needed
   - Adjust the `SLEEP_*` constants in the script if your machine is slower/faster

## Troubleshooting

- If the script fails to interact with the IPOT application, ensure the window title matches `APP_TITLE` in the configuration
- The script uses image recognition and window automation, so the IPOT application must be visible and not minimized
- Make sure all required folders exist before running

## Safety

This script automates mouse and keyboard interactions. Ensure no other important applications are running during execution to avoid accidental interactions.