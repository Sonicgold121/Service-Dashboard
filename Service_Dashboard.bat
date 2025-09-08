@echo off
REM This batch file will start the Email Data Monitoring Dashboard

REM Navigate to the directory where your Streamlit script is located
cd "c:\Users\mwessels\Desktop\Dashboard\"

REM Activate your Python virtual environment if you are using one
REM For example, if your venv is in a folder named "venv" inside the Dashboard directory:
REM call venv\Scripts\activate

REM Run the Streamlit application
echo Starting Streamlit dashboard...
streamlit run "Service Dashboard.py"

REM Optional: Pause to see any messages if Streamlit exits immediately (for troubleshooting)
REM pause