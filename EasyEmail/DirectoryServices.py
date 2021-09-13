"""
Class to do everything that's needed with folders and directories
by: Easton Seidel
2/23/21
"""

from os import listdir, path
from shutil import move, Error


class DirectoryServices:
    """Class to handle the directory services for EasyEmail and Easy Groups"""

    # Declare Directories
    delivered_dir = "delivered_dir"
    undelivered_dir = "undelivered_dir"

    def delivered(self):
        """Method to return the delivered directory"""
        return self.delivered_dir

    def undelivered(self):
        """Method to return the undelivered directory"""
        return self.undelivered_dir

    def get_names(self):
        """Method to get the names of the folders in the undelivered directory"""
        return [f for f in listdir(self.undelivered_dir) if
                path.isdir(path.join(self.undelivered_dir, f))]

    def create_directories(self, names):
        """Method to create a list of the full directories based on the array of names"""
        # Create a new list
        directories = []

        # iterate through the names and create a directory list
        for i in names:
            directories.append(self.undelivered_dir + i)

        return directories

    def move_folder(self, name):
        """Method to move the folders from undelivered to delivered"""
        # Try to move the folder
        try:
            move(self.undelivered_dir + name, self.delivered)
        except Error:
            print("Folder already exists! Saving with .new!")
            move(self.undelivered_dir + name, self.undelivered_dir + name + ".new")
            move(self.undelivered_dir + name + ".new", self.delivered_dir)
