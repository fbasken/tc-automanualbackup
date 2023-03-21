import ctypes
import os
import sys
import shutil
from datetime import datetime
from os.path import expanduser

# C drive special folders
# User special folders


REQUIRE_ADMIN = False

# Backup root folder name
backupFolderName = "_BACKUPS"

# Where to backup
backupFolderPath = os.path.join(
    expanduser("~"), "Desktop", backupFolderName)

# List of paths to back up
folderPathsToCopy = [
    os.path.join(
        "AppData", "Local", "Microsoft", "Outlook", "Offline Address Books"),
    os.path.join(
        "AppData", "Local", "Microsoft", "Edge", "User Data", "Default"),
    # os.path.join(
    # "AppData", "Local", "Google", "Chrome", "User Data", "Default"),
    os.path.join(
        "AppData", "Local", "Mozilla", "Firefox", "Profiles"),
    os.path.join(
        "AppData", "Roaming", "Microsoft", "Templates"),
    os.path.join(
        "AppData", "Roaming", "Microsoft", "Signatures"),
    os.path.join(
        "AppData", "Roaming", "Microsoft", "UProof"),
    os.path.join(
        "AppData", "Roaming", "Microsoft", "Proof"),
    os.path.join(
        "AppData", "Roaming", "Mozilla", "Firefox", "Profiles"),
    os.path.join(
        "Contact"),
    os.path.join(
        "Desktop"),
    os.path.join(
        "Documents"),
    os.path.join(
        "Downloads"),
    os.path.join(
        "Favorites"),
    os.path.join(
        "Music"),
    os.path.join(
        "Pictures"),
    os.path.join(
        "Saved Games"),
    os.path.join(
        "Videos"),
    os.path.join(
        "Work Folders")
]



# https://stackoverflow.com/a/1026626
# Returns true if application is running with elevated privileges
def isAdmin():
    try:
        is_admin = (os.getuid() == 0)
    except AttributeError:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
    return is_admin

# Main program method


def main():

    print("Backup path:\n" + backupFolderPath + "\n")
    # Check if we already have a backup folder
    while (os.path.exists(backupFolderPath)):
        print("Backup folder already exists. If you would like to perform a backup, please delete/move/rename the existing backup folder.")
        # print("Backup folder already exists. Would you like to make a new backup folder? ['y' to continue, anything else to exit] ")
        input("Press Enter to retry\n")

    while True:

        validUsername = False
        while (not validUsername):
            username = input("Enter username to backup (or type 'exit'): ")

            if (username.lower().strip() == "exit"):
                # End program
                sys.exit()

            userHomeFolder = getHomeFolderForUsername(username)
            validUsername = os.path.exists(userHomeFolder)

            if (not validUsername):
                print("User does not exist.\n")
            else:
                lastModifiedDate = datetime.fromtimestamp(os.path.getmtime(userHomeFolder))
                timeSinceModified = datetime.now() - lastModifiedDate

                monthsSinceModified = timeSinceModified.days / 30.4167

                print("Last modified: " + str(round(monthsSinceModified, 1)) + " months ago")
                if (monthsSinceModified > 3):
                    print("Skipping backup for " + username + "; it has been over 3 months since this user folder was modified.")
                    validUsername = False


        print()
        backupUser(username)
        print()


def backupUser(username):
    userHomeFolder = getHomeFolderForUsername(username)

    # Back up each listed folder path
    for relativePath in folderPathsToCopy:
        src_path = os.path.join(userHomeFolder, relativePath)
        dest_path = os.path.join(backupFolderPath, username, relativePath)

        if (os.path.exists(src_path)):
            print("Copying path  | " + src_path)
            # shutil.copytree(src_path, dest_path)
        else:
            print("Skipping path | " + src_path + " | (Folder not found)")
    
    # Back up .PST files
    src_path = os.path.join(userHomeFolder, "AppData", "Local", "Microsoft", "Outlook")
    src_path = os.path.join(userHomeFolder, "Desktop")
    dest_path = os.path.join(backupFolderPath, username, "AppData", "Local", "Microsoft", "Outlook")


    for filename in os.listdir(src_path):
        if filename.endswith(".pst"):
            fullFilePath = os.path.join(src_path, filename)
            os.makedirs(dest_path, exist_ok= True)
            
            print("Copying file  | " + fullFilePath)
            shutil.copy(fullFilePath, dest_path)

    # Hide AppData folder
    ctypes.windll.kernel32.SetFileAttributesW(expanduser("~") + "\\Desktop\\_BACKUPS\\" + username + "\\AppData", 2)


# Returns the path to the user's home folder
def getHomeFolderForUsername(username):
    return os.path.join( "\\", "Users", username)

# -------------------------------------------------------------------- #

if not REQUIRE_ADMIN or isAdmin():
    main()
else:
    # https://stackoverflow.com/a/41930586
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, " ".join(sys.argv), None, 1)

