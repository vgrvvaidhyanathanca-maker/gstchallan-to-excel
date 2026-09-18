@echo off
:: Double-click to tabulate every GST challan PDF in the challan folder.
:: Drop new challan PDFs into CHALLAN_DIR, run this, open GST_Challans.xlsx.
set CHALLAN_DIR=C:\Users\one\Downloads\challan

echo Reading GST challans from %CHALLAN_DIR% ...
echo.
py "%~dp0gst_challan_extract.py" "%CHALLAN_DIR%"
echo.
if errorlevel 1 (
    echo Something went wrong - see the messages above.
) else (
    echo Done. Opening the Excel file...
    start "" "%CHALLAN_DIR%\GST_Challans.xlsx"
)
echo.
pause
