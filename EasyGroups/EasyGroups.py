"""
Just a class
A lovely class to add additional groups
By: Me, Easton Seidel
"""

import io
from datetime import datetime
from time import sleep
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from smtplib import SMTP
import pyad
from pyad import pyad
import keyring
from EasyEmail.DirectoryServices import DirectoryServices


class EasyGroups:
    """Class to handle new hire profiles and their respective groups"""

    def __init__(self, q):
        """ Initialization for the class """
        # Query and other AD info
        self.q = q
        self.domain_name = "domain"
        self.top_level_domain = "local"

        # Initialize AD
        pyad.set_defaults(ldap_server="domain.local", username="domain\svcPDFFiller",
                          password=keyring.get_password("system", "svcDocxFiller"))

        # Initialize Directory Services
        self.ds = DirectoryServices()

        # Variable for who was done and who was done previously
        self.previous = []
        self.done = []
        self.fail = []

        # This rounds new hires
        self.this_round = []

        # Get old log data
        self.old_log()

        # Date Variables
        now = datetime.now()
        self.date = now.strftime("%m-%d-%y")

        # Initialize group values
        self.init_groups()

    def add_groups(self):
        """ Method to add AD groups to new hires """
        # Print a statement to show a change in program
        print("\n* * * * * * * * * * * * * * * * * * * * * * * *"
              "\n* * * * * * * * * Easy Groups * * * * * * * * *"
              "\n* * * * * * * * * * * * * * * * * * * * * * * *")

        # Get the new hires
        users_og = self.ds.get_names()
        users = []

        # Filter out users that already have their groups
        for i in users_og:
            if i not in self.previous:
                users.append(i)

        # Print a statement if list is empty
        if len(users) == 0:
            print("\nThe query returned no new users.")

        # Loop through the users and add the groups
        for i in users:
            # Wait a little bit of time
            sleep(2)

            # Assign the user
            user = pyad.aduser.ADUser.from_cn(i)

            # Query for the users title
            try:
                self.q.execute_query(
                    attributes=["description", "distinguishedName"],
                    where_clause="displayName = '{}'".format(i),
                    base_dn=f"DC={self.domain_name}, DC={self.top_level_domain}"
                )
            except Exception as error:
                print(f"Error with query! {error}")

            # Break if the user doesn't exist anymore
            if self.q.get_row_count() <= 0:
                print(i + " does not exist on " + self.domain_name + "." + self.top_level_domain +
                      "\nPlease verify that the user is still scheduled for onboarding.")
            else:
                # Separate groups
                for row in self.q.get_results():
                    title = row["description"]

                    # Get the user location
                    user_loc = str(row['distinguishedName']).replace(",OU=", ",DC=")
                    user_loc = user_loc.split(',DC=')
                    user_loc = user_loc[2]

                # Convert title to a string
                title = title[0]

                # Add to archive if not corp
                if user_loc != 'Corporate - 1001 (Salt Lake City\\, UT)':
                    user.add_to_group(self.archive)

                # Make sure the user hasn't already been done
                if i not in self.previous:
                    # Try this and print an error otherwise
                    try:
                        # Loop through the list with the matching title
                        if title in ["Loan Officer", "Loan Officer Assistant", "Loan Coordinator",
                                     "Production Coordinator", "Loan Officer Trainee"]:
                            # Add to groups
                            user.add_to_group(self.lo)
                            user.add_to_group(self.production)
                            user.add_to_group(self.non_exempt)

                            # Message for preset
                            message = f"User {i} has been updated using Loan Officer/Coordinator " \
                                      f"and Production Coordinator preset."

                            # Print out what was done
                            print(message)

                            # Append to done list
                            self.done.append(i)
                        elif title == "Loan Processor":
                            # Add to groups
                            user.add_to_group(self.lo)
                            user.add_to_group(self.non_exempt)
                            user.add_to_group(self.processor)
                            user.add_to_group(self.production)

                            # Message for preset
                            message = f"User {i} has been updated using Loan Processor preset."

                            # Print out what was done
                            print(message)

                            # Append to done list
                            self.done.append(i)
                        elif title == "Branch Manager":
                            # Add to groups
                            user.add_to_group(self.bm)
                            user.add_to_group(self.exempt)
                            user.add_to_group(self.receptionist)
                            user.add_to_group(self.hdrive)
                            user.add_to_group(self.lo)
                            user.add_to_group(self.production)

                            # Message for preset
                            message = f"User {i} has been updated using Branch Manager preset."

                            # Print out what was done
                            print(message)

                            # Append to done list
                            self.done.append(i)
                        elif title == "Receptionist":
                            # Add to groups
                            user.add_to_group(self.non_exempt)
                            user.add_to_group(self.production)

                            # Message for preset
                            message = f"User {i} has been updated using Receptionist preset."

                            # Print out what was done
                            print(message)

                            # Append to done list
                            self.done.append(i)
                        elif title in ["Closer", "Insuring Specialist"]:
                            # Add to groups
                            user.add_to_group(self.cie)
                            user.add_to_group(self.foxit)
                            user.add_to_group(self.ops)
                            user.add_to_group(self.ops_fs)
                            user.add_to_group(self.hdrive)
                            user.add_to_group(self.sharefile)

                            # Message for preset
                            message = f"User {i} has been updated using a Closing preset."

                            # Print out what was done
                            print(message)

                            # Append to done list
                            self.done.append(i)
                        elif title in ["Funder", "Funding Assistant"]:
                            # Add to groups
                            user.add_to_group(self.foxit)
                            user.add_to_group(self.ops)
                            user.add_to_group(self.ops_fs)
                            user.add_to_group(self.hdrive)
                            user.add_to_group(self.sharefile)

                            # Message for preset
                            message = f"User {i} has been updated using a Funder preset."

                            # Print out what was done
                            print(message)

                            # Append to done list
                            self.done.append(i)
                        elif title in ["Final Docs Specialist", "Shipper",
                                       "Post-Closing Specialist", "Post-Closing Specialist I",
                                       "Post-Closing Specialist II", "Post-Closing Specialist III",
                                       "Post Closing Specialist", "Post Closing Specialist I",
                                       "Post Closing Specialist II", "Post Closing Specialist III"]:
                            # Add to groups
                            user.add_to_group(self.foxit)
                            user.add_to_group(self.filezilla)
                            user.add_to_group(self.ops)
                            user.add_to_group(self.ops_fs)
                            user.add_to_group(self.hdrive)
                            user.add_to_group(self.sharefile)

                            # Message for preset
                            message = f"User {i} has been updated using an Operations preset."

                            # Print out what was done
                            print(message)

                            # Append to done list
                            self.done.append(i)
                        elif title == "Collateral Specialist":
                            # Add to groups
                            user.add_to_group(self.foxit)
                            user.add_to_group(self.filezilla)

                            # Message for Preset
                            message = f"User {i} has been updated using Collateral Specialist " \
                                      f"preset."

                            # Print out what was done
                            print(message)

                            # Append to done list
                            self.done.append(i)
                        else:
                            # Message for Preset
                            message = f"User {i} did not match any presets"

                            # Print out what was done
                            print(message)

                            self.fail.append(i)
                    except Exception as error:
                        message = f"There was an error with adding groups for {i}: {error}!"
                        print(message)
                else:
                    message = f"User {i} already has their groups."
                    print(message)

            self.run_log(message)

            # Add the user to this round's modified new hires
            self.this_round.append(message)

        # Write the log
        self.finished_logs()

        # Send the email
        if len(self.this_round) != 0:
            self.email_confirmation()

    def run_log(self, message):
        """Method to create a new run log"""
        with io.open(f"H:\\New Hire Onboarding\\Program Logs\\EasyGroups\\Logs\\{self.date}.txt",
                     "a") as log: log.write(message + "\n")

    def finished_logs(self):
        """Method to write out our logs"""
        # Open the log file
        with io.open("H:\\New Hire Onboarding\\Program Logs\\EasyGroups\\Finished.txt", "a") as log:
            # Loop through the users and add them
            for i in self.done:
                log.write(f"\n{i}")

        # Create the failed log
        with io.open("H:\\New Hire Onboarding\\Program Logs\\EasyGroups\\Failed.txt", "a") as log:
            # Loop through the users and add them
            for i in self.fail:
                log.write(f"\n{i}")

    def old_log(self):
        """Method to import the old log"""
        # open the log file and read all the lines
        with io.open("H:\\New Hire Onboarding\\Program Logs\\EasyGroups\\Finished.txt", "r") as log:
            for line in log:
                self.previous.append(line.rstrip('\n'))

    def email_confirmation(self):
        """Method for sending what was done"""

        msg = MIMEMultipart('related')
        msg['From'] = "EasyGroups@domain.com"
        msg['Subject'] = "New Hire Groups"

        if len(self.this_round) != 0:
            breaks = "<br><br>"
        else:
            breaks = ""

        new_line = '<br>'
        msg_body = MIMEText("<p>Hello!</p>"
                            "<p>The following changes have been made to new hires:</p>"
                            f"<p>{(new_line.join(self.this_round))}{breaks}"
                            "<br>Please report any strange behavior to "
                            "<strong>Easton Seidel</strong>.</p>"
                            "<p>Thanks,<br>svcDocxFiller</p>", "html")
        msg.attach(msg_body)

        smtp_server = SMTP("smtp_ip:25")

        try:
            receiver = "domain-servicedesk@domain.com"
            # receiver = "eseidel@domain.com"
            msg['To'] = receiver
            smtp_server.sendmail(msg["From"], receiver, msg.as_string())
        except Exception as error:
            print(f"Failed to send post run email: {error}")

        smtp_server.quit()

    def init_groups(self):
        """Method to house all of our AD groups"""
        # domain-MailArchive
        self.archive = pyad.adgroup.ADGroup.from_dn("CN=domain-MailArchive,OU=Security Groups,"
                                                    "OU=Exchange Objects,OU=Offices,"
                                                    "DC=domain,DC=local")
        # domain-LoanOfficers
        self.lo = pyad.adgroup.ADGroup.from_dn("CN=domain-Loan Officers,OU=Distribution Group,"
                                               "OU=Exchange Objects,OU=Offices,"
                                               "DC=domain,DC=local")
        # domain-ProductionUsers
        self.production = pyad.adgroup.ADGroup.from_dn("CN=domain-Production Users,OU=Citrix,"
                                                       "OU=Groups,DC=domain,DC=local")
        # domain-Non-ExemptEmployee
        self.non_exempt = pyad.adgroup.ADGroup.from_dn("CN=domain-Non-Exempt Employee,"
                                                       "OU=Distribution Group,OU=Exchange Objects,"
                                                       "OU=Offices,DC=domain,DC=local")
        # domain-Processors
        self.processor = pyad.adgroup.ADGroup.from_dn("CN=domain-Processors,OU=Distribution Group,"
                                                      "OU=Exchange Objects,OU=Offices,"
                                                      "DC=domain,DC=local")
        # domain-Branch Managers
        self.bm = pyad.adgroup.ADGroup.from_dn("CN=domain-Branch Managers,OU=Distribution Group,"
                                               "OU=Exchange Objects,OU=Offices,"
                                               "DC=domain,DC=local")
        # domain-Exempt
        self.exempt = pyad.adgroup.ADGroup.from_dn("CN=domain-Exempt Employee,OU=Distribution Group,"
                                                   "OU=Exchange Objects,OU=Offices,"
                                                   "DC=domain,DC=local")
        # domain-FS-Receptionist
        self.receptionist = pyad.adgroup.ADGroup.from_dn("CN=domain-FS-Reception,OU=File Server,"
                                                         "OU=Groups,DC=domain,"
                                                         "DC=local")
        # domain-FS-HDriveAccess
        self.hdrive = pyad.adgroup.ADGroup.from_dn("CN=domain-FS-HDriveAccess,OU=File Server,"
                                                   "OU=Groups,DC=domain,DC=local")
        # domain-CIE
        self.cie = pyad.adgroup.ADGroup.from_dn("CN=domain-CIE,OU=Distribution Group,"
                                                "OU=Exchange Objects,OU=Offices,"
                                                "DC=domain,DC=local")
        # domain-Foxit Users
        self.foxit = pyad.adgroup.ADGroup.from_dn("CN=domain-Foxit Users,OU=Citrix,OU=Groups,"
                                                  "DC=domain,DC=local")
        # domain-Ops Department
        self.ops = pyad.adgroup.ADGroup.from_dn("CN=domain-OPS,OU=RDS,OU=Groups,"
                                                "DC=domain,DC=local")

        # domain-RDFilezilla
        self.filezilla = pyad.adgroup.ADGroup.from_dn("CN=domain-RDFilezilla,OU=RDS,OU=Groups,"
                                                      "DC=domain,DC=local")
        # Ops_Department
        self.ops_fs = pyad.adgroup.ADGroup.from_dn("CN=domain-FS-Ops_Department,OU=File Server,"
                                                   "OU=Groups,DC=domain,DC=local")
        # Sharefile
        self.sharefile = pyad.adgroup.ADGroup.from_dn("CN=domain-Sharefile,OU=Citrix,OU=Groups,"
                                                      "DC=domain,DC=local")
