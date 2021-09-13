"""
Main file for all the automation control
By: Easton Seidel and Ethan Berg

Compile Command pyinstaller --hidden-import pywin32 --hidden-import win32timezone --onefile main.py
"""

import time
from datetime import datetime
import pyad.adquery
import EasyEmail.SendEmails as se
from EasyGroups.EasyGroups import EasyGroups
import letters.PyLetters


q = pyad.adquery.ADQuery()
eg = EasyGroups(q)


def main():
    """Main driver for the automatatron new hire provisioning software"""
    # Get the time values
    today = datetime.now()
    hour = int(today.strftime("%H"))

    # Get the day of the week
    dow = datetime.today().weekday()   # Days start at 0 with monday and end with 6 as sunday

    # Email Launcher
    if dow == 4 and hour == 16:
        time.sleep(360)
        se.send_emails(q)

    # Run these every time
    letters.PyLetters.create_letters(q)
    time.sleep(10)
    eg.add_groups()


if __name__ == '__main__':
    main()
