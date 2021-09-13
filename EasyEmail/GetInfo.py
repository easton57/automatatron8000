"""
Class to get the required info from the user
by: Easton Seidel
2/23/21
"""

from getpass import getpass


class GetInfo:
    """Method to get email and username for sending emails via outlook"""
    domain = "@domain.com"

    def get_email(self):
        """Method to get the users email when manually authenticating"""
        print("\nThe rest of the email domain will be added, no need to enter it.")

        while True:
            print("Please enter your username: ", end='')
            email = input() + self.domain

            # Verify the email with the user
            print("\n" + email + "\n")
            print("Is the above email correct? Enter YES: ", end='')

            if input() == "YES":
                return email

    @staticmethod
    def get_password():
        """Method to get the users password when manually authenticating"""
        while True:
            password = getpass(prompt="Please enter your password: ")

            confirmation = getpass(prompt="Please confirm your password: ")

            if confirmation == password:
                return password

            print("Passwords do not match, please try again.")
