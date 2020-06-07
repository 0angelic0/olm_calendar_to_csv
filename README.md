# olm_calendar_to_csv
Extract Outlook for Macs calendars (appointments) olm file to csv file.

# Requirement
You have to have python3.7 in your mac.

# How to use
1. From Outlook for Macs export only calendars (appointments) to olm file.
2. unzip the olm file.
3. Locate the Calendar.xml file (Accounts/{your_email}/Calendar/Calendar.xml).
4. Place extract_calendar.py in the same directory of Calendar.xml.
5. Open terminal then cd to the same directory of Calendar.xml and extract_calendar.py.
6. Run command in your terminal.

`python extract_calendar.py`

6. You will find new `output.csv` file.
7. Done.
8. (Optionally) Import the `output.csv` to Excel for Macs by using Tab as a seperator.

# Need Comma instead of Tab
Just change the dialect in extract_calendar.py at line 34

from `excel-tab` to `excel`. Then you're good to go.
