@echo off
:: Navigate to the directory where the script is located
if exist "C:\Automation\Automated_Scheduled_Jobs\Python_TRCM-DeskTickect_Activity_Reporting" (
    cd /d "C:\Automation\Automated_Scheduled_Jobs\Python_TRCM-DeskTickect_Activity_Reporting"
) else (
    echo Directory not found: C:\Automation\Automated_Scheduled_Jobs\Python_TRCM-DeskTickect_Activity_Reporting
    pause
    exit /b
)

:: Check if the script exists
if exist "TrueRCM_Desk_Tickets_Activity_Reporting_ScrollAdded_Streamlit_v4.py" (
    python "TrueRCM_Desk_Tickets_Activity_Reporting_ScrollAdded_Streamlit_v4.py"
) else (
    echo Script file not found: TrueRCM_Desk_Tickets_Activity_Reporting_ScrollAdded_Streamlit_v4.py
)
pause
