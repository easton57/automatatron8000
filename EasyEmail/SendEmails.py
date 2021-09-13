"""
Program to email the new hire letters
By: Easton Seidel
2/23/21
"""

from EasyEmail.EmailServices import EmailServices
from EasyEmail.ConfirmedNewHires import ConfirmedNewHires
from EasyEmail.DirectoryServices import DirectoryServices
from EasyEmail.GetInfo import GetInfo

ds = DirectoryServices()
cnh = ConfirmedNewHires()
gi = GetInfo()


def send_emails(q):
    """Gather needed information and send emails to the new hires/managers"""
    # Get the names of the users
    # names = ds.get_names()

    # Verify which users need the email
    # names = cnh.confirm_letters(names)
    names = cnh.auto_confirm()

    # Quit if list is empty
    if names == 0:
        exit()

    # Get the username and password of the user
    # email = gi.get_email()
    # password = gi.get_password()

    # Indicate a change in program
    print("\n* * * * * * * * * * * * * * * * * * * * * * * *"
          "\n* * * * * * * * * Easy Email  * * * * * * * * *"
          "\n* * * * * * * * * * * * * * * * * * * * * * * *")

    # Send the letters and move the folders
    for i in names:

        # initialize the email services
        es = EmailServices(q)

        # Send the emails
        es.send_email(i)

        # move the folders
        try:
            ds.move_folder(i)
        except Exception:
            print(f"{i}.new folder exists! Clean up a bit and move the folder!")
