"""
Class for verifying which users need the emails this week.
by: Easton Seidel
2/23/21
"""

import pandas as pd


class ConfirmedNewHires:
    """Class to automatically or manually confirm new hires"""

    def confirm_letters(self, old_letters):
        """Method to print and select confirmed letters from the undelivered directory"""
        # loop and make sure we get numbers
        while True:
            # Print the list of the letters
            print("These are the letters that have been found.\n")

            # Function to print them all
            self.print_list(old_letters)

            print("\nWhich new hires will be starting next week?")
            print("Separate numbers with a space: ", end='')

            # Capture their input
            numbers = input()
            new_numbers = []

            # Needed for double digits
            prev_space = False
            j = 0

            # Create an array with the numbers that were entered
            for i in numbers:
                if i == ' ' and prev_space:
                    prev_space = False
                elif prev_space:
                    j -= 1
                    new_numbers[j] += i
                else:
                    new_numbers.append(i)
                    prev_space = True

                # Make sure all the numbers are within range and remove a number if it is outside
                # of the range
                if int(new_numbers[j]) >= len(old_letters):
                    print("You entered a number outside of the range of the list.")
                    new_numbers.remove(new_numbers[j])

                j += 1

            # Pull the confirmed new hires
            confirmed = []

            for i in new_numbers:
                confirmed.append(old_letters[int(i)])

            # Check to see if the list is empty and see if the user wants to quit
            if not confirmed:
                print("Your list is empty. Would you like to exit the program? Enter YES: ", end='')
                confirmation = input()
                if confirmation == "YES":
                    return 0

            print("\nThese are your confirmed users:\n")
            self.print_list(confirmed)

            finished = input("\nIs the above list correct? Enter YES: ")
            if finished == "YES":
                break

        return confirmed

    @staticmethod
    def auto_confirm():
        """Method to check a csv and select only the names listed in that csv"""
        starting_next_week = "new hire csv directory"
        confirmed = []

        # Pull the confirmed users from the csv
        data = pd.read_csv(starting_next_week)

        for i in range(len(data)):
            confirmed.append(data.loc[i][1])

        return confirmed

    @staticmethod
    def print_list(users):
        """Method to print the names of users from the provided array"""
        for i in range(len(users)):
            print(str(i) + ": " + users[i])
