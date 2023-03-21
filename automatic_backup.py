import ctypes
import os
import sys
import shutil
from datetime import datetime
from os.path import expanduser

# TODO:
# Change backup location to C:\ITS\_BACKUP


REQUIRE_ADMIN = False

# Backup root folder name
backupFolderName = "_BACKUP"

# Where to backup
backupFolderPath = os.path.join(
    expanduser("~"), "Desktop", backupFolderName)

# List of paths to back up
folderPathsToCopy = [
    os.path.join(
        "AppData", "Local", "Microsoft", "Outlook", "Offline Address Books"),
    os.path.join(
        "AppData", "Local", "Microsoft", "Edge", "User Data", "Default"),
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
        "Contacts"),
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

standardItemsInUserFolder = [
    ".cisco",
    ".ms-ad",
    "3D Objects",
    "Contacts",
    "Desktop",
    "Documents",
    "Downloads",
    "Favorites",
    "Links",
    "Music",
    "Pictures",
    "Saved Games",
    "Searches",
    "Videos",
    "Work Folders"
]

standardItemsInRootFolder = [
    "$WinREAgent",
    "APPDIR",
    "Dell",
    "Intel",
    "ITS",
    "MSOCache",
    "PerfLogs",
    "Program Files",
    "Program Files (x86)",
    "ProgramData",
    "temp",
    "Users",
    "Windows",
    "Config.Msi",
    "Documents and Settings",
    "Recovery",
    "System Volume Information"
]

# Prevent OS sleep/hibernate in windows; code from:
# https://github.com/h3llrais3r/Deluge-PreventSuspendPlus/blob/master/preventsuspendplus/core.py
# API documentation:
# https://msdn.microsoft.com/en-us/library/windows/desktop/aa373208(v=vs.85).aspx
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001

def enableWindowsSleepPrevention(preventSleep):
    if (preventSleep):
        print("\n-- Preventing Windows from going to sleep")
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED)
    else:
        print("\n-- Allowing Windows to go to sleep")
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)

# https://stackoverflow.com/a/1026626
# Returns true if application is running with elevated privileges
def isAdmin():
    try:
        is_admin = (os.getuid() == 0)
    except AttributeError:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
    return is_admin

# Main program method

def isHidden(path):
    return bool(ctypes.windll.kernel32.GetFileAttributesW(path) & 2)

def displayHeader(headerText):
    print("\n-- " + headerText)
    
def main():

    print("Backup path:\n" + backupFolderPath + "\n")
    # Check if we already have a backup folder
    while (os.path.exists(backupFolderPath)):
        print("Backup folder already exists. If you would like to perform a backup, please delete/move/rename the existing backup folder.")
        # print("Backup folder already exists. Would you like to make a new backup folder? ['y' to continue, anything else to exit] ")
        input("Press Enter to retry.\n")

    print()
    backupRoot()
    print()

    while True:

        validUsername = False
        while (not validUsername):
            username = input("Enter username to backup (or type 'exit'): ")

            if (username.lower().strip() == "exit"):
                # End program
                input("\n--\nBackup folder is on your Desktop.\nPress Enter to quit.")
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

# Given a source path, and a destination folder, will copy the source file/folder (and subfolders)
# and create a destination path if necessary
def safeCopy(src_path, dest_path):
    paddedOutput = src_path.ljust(40)

    # Check if the file/folder to copy even exists
    if (os.path.exists(src_path)):

        # If the SOURCE path to copy is a FILE, copy it as a file
        if os.path.isfile(src_path):
            # Then, if the destination doesn't exist, make it
            if(not os.path.exists(dest_path)):
                os.makedirs(dest_path, exist_ok= True)

            print("Copying file  | " + paddedOutput)
            shutil.copy(src_path, dest_path)

        # If the SOURCE path to copy is a FOLDER, copy it recursively
        else:
            # Then, if the destination DOES exist, this will FAIL
            if(os.path.exists(dest_path)):
                # This should not happen ever
                print("Failed !!!    | " + paddedOutput + " | (Did not copy folder since destination already exists)")
            else:
                print("Copying dir   | " + paddedOutput)
                shutil.copytree(src_path, dest_path)
    else:
        print("Skipping path | " + paddedOutput + " | (Does not exist)")

def backupRoot():
    print("[ BACKING UP ROOT FOLDERS (Non-user specific) ]")
    
    enableWindowsSleepPrevention(True)

    # Back up C:\SAS
    displayHeader("Backing up C:\\SAS\\")
    src_path = "\\SAS"
    dest_path = backupFolderPath

    safeCopy(src_path, dest_path)

    # Back up all non-standard folders on C:
    displayHeader("Backing up non-standard folders on C:\\")
    src_path = "\\"
    dest_path = backupFolderPath

    for filename in os.listdir("\\"):
            fullFilePath = src_path + filename
            paddedOutput = fullFilePath.ljust(40)
            # Skip folders for these reasons:

            # If it's not a folder (it's a file)
            if os.path.isfile(fullFilePath):
                print("Skipping path | " + paddedOutput + " | (Is a file, not a folder)")
                continue
            # If the folder is standard
            if filename in standardItemsInRootFolder:
                print("Skipping path | " + paddedOutput + " | (Standard folder)")
                continue
            # If the folder starts with . or $
            if filename.startswith(".") or filename.startswith("$"):
                print("Skipping path | " + paddedOutput + " | (Skipped due to prefix)")
                continue
            # If the folder is hidden
            if(isHidden(fullFilePath)):
                print("Skipping path | " + paddedOutput + " | (Hidden folder)")
                continue
            
            safeCopy(fullFilePath, os.path.join(dest_path, filename))

    enableWindowsSleepPrevention(False)

    print("\n[DONE]")

def backupUser(username):
    print("[ BACKING UP USER: '" + username + "' ]")

    enableWindowsSleepPrevention(True)

    userHomeFolder = getHomeFolderForUsername(username)
    backupUsersFolderPath = os.path.join(backupFolderPath, "Users")

    # Back up each listed folder path
    displayHeader("Backing up user folders and data")
    for relativePath in folderPathsToCopy:
        src_path = os.path.join(userHomeFolder, relativePath)
        dest_path = os.path.join(backupUsersFolderPath, username, relativePath)

        # safeCopy(src_path, dest_path)
    
    # Back up .PST files
    src_path = os.path.join(userHomeFolder, "AppData", "Local", "Microsoft", "Outlook")
    src_path = os.path.join(userHomeFolder, "Desktop") # DELETE THIS
    dest_path = os.path.join(backupUsersFolderPath, username, "AppData", "Local", "Microsoft", "Outlook")

    displayHeader("Backing up Microsoft Outlook .pst files")
    for filename in os.listdir(src_path):
        if filename.endswith(".pst"):
            fullFilePath = os.path.join(src_path, filename)
            
            safeCopy(fullFilePath, dest_path)

    # Back up Chrome\User Data\Default, except for Service Worker folder
    src_path = os.path.join(userHomeFolder, "AppData", "Local", "Google", "Chrome", "User Data", "Default")
    dest_path = os.path.join(backupUsersFolderPath, username, "AppData", "Local", "Google", "Chrome", "User Data", "Default")

    if os.path.exists(src_path):
        displayHeader("Backing up Chrome User Data files (except Service Worker folder)")

        for filename in os.listdir(src_path):
            fullFilePath = os.path.join(src_path, filename)

            # If we find a folder named 'Service Worker'
            if filename == "Service Worker" and not os.path.isfile(fullFilePath):
                # Skip the folder
                continue
            
            safeCopy(fullFilePath, dest_path)

    # Back up all non-standard folders in user folder
    src_path = userHomeFolder
    dest_path = os.path.join(backupUsersFolderPath, username)

    if os.path.exists(src_path):
        displayHeader("Backing up non-standard folders in '" + username + "' user folder")

        for filename in os.listdir(src_path):
            fullFilePath = os.path.join(src_path, filename)
            paddedOutput = fullFilePath.ljust(40)

            # Skip certain folders for these reasons:

            # If it's not a folder (it's a file)
            if os.path.isfile(fullFilePath):
                print("Skipping path | " + paddedOutput + " | (Is a file, not a folder)")
                continue
            # If the folder is standard
            if filename in standardItemsInUserFolder:
                print("Skipping path | " + paddedOutput + " | (Standard folder)")
                continue
            # If the folder starts with . or $
            if filename.startswith(".") or filename.startswith("$"):
                print("Skipping path | " + paddedOutput + " | (Skipped due to prefix)")
                continue
            # If the folder is hidden
            if(isHidden(fullFilePath)):
                print("Skipping path | " + paddedOutput + " | (Hidden folder)")
                continue

            safeCopy(fullFilePath, os.path.join(dest_path, filename))

    # Hide AppData folder
    appdataPath = os.path.join(backupUsersFolderPath, username, "AppData")
    displayHeader("Marking AppData folder as hidden " + appdataPath)
    if os.path.exists(appdataPath):
        ctypes.windll.kernel32.SetFileAttributesW(appdataPath, 2)
        print("AppData folder for '" + username + "' successfully hidden.")
    else:
        print("Error! Could not hide AppData folder-- AppData folder for '" + username + "' does not exist.")
    
    enableWindowsSleepPrevention(False)

    print("\n[DONE]")

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

