"""
Class for the email services needed to send the emails
Written to send from local outlook or SMTP.
Set SMTP to false to send from local outlook.
by: Easton Seidel
2/23/21
"""

import io
from smtplib import SMTP
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from datetime import datetime
from EasyEmail.DirectoryServices import DirectoryServices

ds = DirectoryServices()


class EmailServices:
    """Class to handle all the email sending"""

    def __init__(self, q):
        self.smtp_ip = "smtp_ip:25"
        self.recipient = ""
        self.display_name = ""
        self.first_name = ""
        self.manager_email = ""
        self.manager = ""
        self.manager_name = ""
        self.manager_first_name = ""
        self.corp_locations = ['']

        # Log Variables
        self.log_dir = "log_dir"

        # Date Variables
        now = datetime.now()
        self.time = now.strftime("%m-%d-%y")
        self.file = io.open(self.log_dir + str(self.time) + ".txt", "a")

        # SMTP Variables
        self.smpt = True
        self.sender = 'it-support@domain.com'
        self.test_recipient = 'eseidel@domain.com'

        # Signature stuff since it won't go through exclaimer
        self.logo = "Signatures\\logo.png"
        self.phone = "Signatures\\phone.png"
        self.letter = "Signatures\\letter.png"

        # initialize the stuff
        # AD Stuff
        self.q = q
        self.domain_name = "domain"
        self.top_level_domain = "local"

    def send_email(self, display_name):
        """Class to gather information and initialize the senders for various emails"""
        # Pull the first name and assign it to self.first_name
        self.display_name = display_name
        self.pull_first_name()

        # Get the users info from AD
        # Taken from Ethans code. Thanks Ethan.
        self.q.execute_query(
            attributes=["mail", "manager", "description", "distinguishedName"],
            where_clause="displayName = '{}'".format(display_name),
            base_dn=f"DC={self.domain_name}, DC={self.top_level_domain}"
        )

        if self.q.get_row_count() <= 0:
            print(display_name + " does not exist on " + self.domain_name + "." +
                  self.top_level_domain + "\nPlease verify that the user is still scheduled "
                                          "for onboarding.")
            return False

        for row in self.q.get_results():
            self.recipient = row["mail"]
            self.manager = row["manager"]
            title = row["description"]

            # Get the user location
            user_loc = str(row['distinguishedName']).replace(",OU=", ",DC=")
            user_loc = user_loc.split(',DC=')
            user_loc = user_loc[2]

        # Call the sender modules
        self.send_employee_smtp()
        self.send_manager_smtp()

        self.file.close()
        return True

    def pull_first_name(self):
        """Method that pulls the first name of the user"""
        for i in self.display_name:

            if i == ' ':
                break

            self.first_name += i

    def pull_manager_name(self):
        """Method to pull the managers name from AD"""
        # assign the manager string to a temp variable
        temp = self.manager

        # pull manager's name from AD query
        cn = True
        for i in temp:
            if i in ('C', 'N', '=') and cn is True:
                if i == '=':
                    cn = False
            elif i == ',':
                break
            else:
                self.manager_name += i

        # pull the first name
        for i in self.manager_name:
            if i == ' ':
                break

            self.manager_first_name += i

        # Grab the managers email
        self.q.execute_query(
            attributes=["mail"],
            where_clause="displayName = '{}'".format(self.manager_name),
            base_dn=f"DC={self.domain_name}, DC={self.top_level_domain}"
        )

        for row in self.q.get_results():
            self.manager_email = row["mail"]

    def send_employee_smtp(self):
        """Method to send the welcome letter to the new hire"""
        file = ds.undelivered() + self.display_name + "\\New Hire Letter " + \
               self.display_name + ".docx"

        msg = MIMEMultipart('related')
        msg['From'] = self.sender
        msg['To'] = self.recipient
        msg['Subject'] = "Login Information"

        # The body of the message in HTML
        msg_text = MIMEText(f"<p>Hello {self.first_name},</p>"
                            f"<p></p>"
                            f"<p>Attached is the login information for your accounts.</p>"
                            f"<p></p>"
                            f"<p>Thanks,</p>"
                            f"<p style='font-size:8.0pt'>Did you know that we have a Frequently "
                            f"Asked Questions and Knowledge "
                            f"Base website? You can also submit tickets there!</p>"
                            f"<p style='font-size:8.0pt'>"
                            f"<a href='https://domain.service-now.com/sp'>"
                            f"https://domain.service-now.com/sp</a></p>"
                            f"<img src='cid:image1'><br>"
                            f"<table><tr align='center'>"
                            f"<th><img src='cid:image2'></th>"
                            f"<th><img src='cid:image3'></th>"
                            f"<th>IT-SUPPORT@domain.com</th>"
                            f"</tr></table>", 'html')
        msg.attach(msg_text)

        # Create signature picture
        sig = io.open(self.logo, 'rb')
        msg_logo = MIMEImage(sig.read())
        sig.close()

        # Create signature picture
        sig = io.open(self.phone, 'rb')
        msg_phone = MIMEImage(sig.read())
        sig.close()

        # Create signature picture
        sig = io.open(self.letter, 'rb')
        msg_letter = MIMEImage(sig.read())
        sig.close()

        # Define the image's ID as referenced above
        msg_logo.add_header('Content-ID', '<image1>')
        msg_phone.add_header('Content-ID', '<image2>')
        msg_letter.add_header('Content-ID', '<image3>')
        msg.attach(msg_logo)
        msg.attach(msg_phone)
        msg.attach(msg_letter)

        # Create and attach letter
        with io.open(file, "rb") as fil:
            part = MIMEApplication(fil.read(), Name=basename(file))

        # After the file is closed
        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(file)
        msg.attach(part)

        smtp_obj = SMTP(self.smtp_ip)

        try:
            smtp_obj.sendmail(msg["From"], msg['To'], msg.as_string())
            status = f"Sent Email to {self.display_name}."
        except Exception as error:
            status = f"Failed to send email to {self.recipient}! {error}"

        # Get the date and time
        now = datetime.now()
        date_time = now.strftime("%m/%d/%Y %H:%M:%S")

        print(status)
        self.file.write(f"{date_time} - {status}\n")

        smtp_obj.quit()

    def send_manager_smtp(self):
        """Method to send the new hire letter to the manager"""
        # Pull the managers name from the employee ad query
        self.pull_manager_name()

        file = ds.undelivered() + self.display_name + "\\New Hire Letter " + self.display_name + \
               " Manager.docx"
        msg = MIMEMultipart('related')
        msg['From'] = self.sender
        msg['To'] = self.manager_email
        msg['Subject'] = "Login Information"

        # The body of the message in HTML
        msg_text = MIMEText(f"<p>Hello {self.manager_first_name},</p>"
                            f"<p></p>"
                            f"<p>Attached is the new hire letter for {self.display_name}.</p>"
                            f"<p></p>"
                            f"<p>Thanks,</p>"
                            f"<p style='font-size:8.0pt'>Did you know that we have a Frequently "
                            f"Asked Questions and Knowledge "
                            f"Base website? You can also submit tickets there!</p>"
                            f"<p style='font-size:8.0pt'>"
                            f"<a href='https://domain.service-now.com/sp'>"
                            f"https://domain.service-now.com/sp</a></p>"
                            f"<img src='cid:image1'><br>"
                            f"<table><tr align='center'>"
                            f"<th><img src='cid:image2'></th>"
                            f"<th><img src='cid:image3'></th>"
                            f"<th>IT-SUPPORT@domain.com</th>"
                            f"</tr></table>", 'html')
        msg.attach(msg_text)

        # Create signature picture
        sig = io.open(self.logo, 'rb')
        msg_logo = MIMEImage(sig.read())
        sig.close()

        # Create signature picture
        sig = io.open(self.phone, 'rb')
        msg_phone = MIMEImage(sig.read())
        sig.close()

        # Create signature picture
        sig = io.open(self.letter, 'rb')
        msg_letter = MIMEImage(sig.read())
        sig.close()

        # Define the image's ID as referenced above
        msg_logo.add_header('Content-ID', '<image1>')
        msg_phone.add_header('Content-ID', '<image2>')
        msg_letter.add_header('Content-ID', '<image3>')
        msg.attach(msg_logo)
        msg.attach(msg_phone)
        msg.attach(msg_letter)

        # Create and attach letter
        with io.open(file, "rb") as fil:
            part = MIMEApplication(fil.read(), Name=basename(file))

        # After the file is closed
        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(file)
        msg.attach(part)

        smtp_obj = SMTP(self.smtp_ip)

        try:
            smtp_obj.sendmail(msg["From"], msg['To'], msg.as_string())
            status = f"Sent Email to {self.manager_name}."
        except Exception as error:
            status = f"Failed to send email to {self.manager_name}! {error}"

        # Get the date and time
        now = datetime.now()
        date_time = now.strftime("%m/%d/%Y %H:%M:%S")

        print(status)
        self.file.write(f"{date_time} - {status}\n")

        smtp_obj.quit()
