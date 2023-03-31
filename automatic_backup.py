'''
AutoManual Backup by Freddie
for DESKTOP side backups

ver 3/30/2023

'''

import logging
import ctypes
import os
import sys
import shutil
from datetime import datetime
from os.path import expanduser

# If True, will attempt to elevate the program on launch (recommended!)
REQUIRE_ADMIN = True

# If True, will disable all actual copying/directory creation. Use when testing.
DEBUG_ONLY_DONT_COPY = False

# Backup root folder name
backupFolderName = "_BACKUP"

# Where to backup
backupFolderPath = "\\ITS\\" + backupFolderName

# List of paths to back up, relative to the user folder
userFolderPathsToCopy = [
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
        "Videos")
]

# List of standard folders found in the user folder
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

# List of root paths to copy
rootFolderPathsToCopy = [
    "\\SAS"
]

# List of standard folders found in the root folder
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

def logPrint(output=""):
    print(output)
    logging.info(output)
    
#   Adapted from https://stackoverflow.com/a/57647169
# Prevent OS sleep/hibernate in windows; code from:
# https://github.com/h3llrais3r/Deluge-PreventSuspendPlus/blob/master/preventsuspendplus/core.py
# API documentation:
# https://msdn.microsoft.com/en-us/library/windows/desktop/aa373208(v=vs.85).aspx
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001

def enableWindowsSleepPrevention(preventSleep):
    if (preventSleep):
        logPrint("\n-- Preventing Windows from going to sleep")
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED)
    else:
        logPrint("\n-- Allowing Windows to go to sleep")
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)

#   From https://stackoverflow.com/a/1026626
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
    logPrint("\n-- " + headerText)
    
def main():
    # Warn if running in debug mode
    if (DEBUG_ONLY_DONT_COPY):
        print("NOTE: The program is running in debug mode. No files will be copied -- this mode is for testing.")
        print("\n\tDEBUG_ONLY_DONT_COPY = True\n\n")


    # Set up logging
    logFileName = "automanualbackup_log" + str(datetime.now().timestamp()) + ".txt"
    logging.basicConfig(filename=logFileName, encoding='utf-8', level=logging.DEBUG)
    print("[Created log file: " + logFileName + "]\n")

    # Inform user of backup folder location
    logPrint("Backup folder path:\n" + backupFolderPath + "\n")

    # Check if we already have a backup folder
    while (os.path.exists(backupFolderPath)):
        logPrint("Backup folder already exists. If you would like to perform a backup, please delete/move/rename the existing backup folder.")
        # logPrint("Backup folder already exists. Would you like to make a new backup folder? [y/N] ")
        input("Press Enter to retry.\n")

    logPrint()
    backupRoot()
    logPrint()

    while True:

        validUsername = False
        while (not validUsername):
            username = input("Enter username to backup (or type 'exit'): ")

            if (username.lower().strip() == "exit"):
                # End program
                logPrint("\n--\n\nBackup folder path:\n" + backupFolderPath + "\n")

                input("Press Enter to quit.")
                
                sys.exit()

            userHomeFolder = getHomeFolderForUsername(username)
            validUsername = os.path.exists(userHomeFolder)

            if (not validUsername):
                logPrint("User does not exist.\n")
            else:
                lastModifiedDate = datetime.fromtimestamp(os.path.getmtime(userHomeFolder))
                timeSinceModified = datetime.now() - lastModifiedDate

                monthsSinceModified = timeSinceModified.days / 30.4167

                logPrint("Last modified: " + str(round(monthsSinceModified, 1)) + " months ago")

                MONTH_LIMIT = 3
                if (monthsSinceModified > MONTH_LIMIT):
                    logPrint("This user folder hasn't been modified in over " + str(MONTH_LIMIT) +" months.")
                    ans = input("Continue with backup? [y/N] ")
                    if ans.lower().strip() == "y":
                        logPrint("Continuing with backup for '" + username + "'\n")
                    else:
                        logPrint("Skipping backup for '" + username + "'\n")
                        validUsername = False


        logPrint()
        backupUser(username)
        logPrint()

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
                if (not DEBUG_ONLY_DONT_COPY):
                    os.makedirs(dest_path, exist_ok= True)

            logPrint("Copying file  | " + paddedOutput)
            if (not DEBUG_ONLY_DONT_COPY):
                shutil.copy(src_path, dest_path)

        # If the SOURCE path to copy is a FOLDER, copy it recursively
        else:
            # Then, if the destination DOES exist, this will FAIL
            if(os.path.exists(dest_path)):
                # This should not happen ever! ! !
                logPrint("Failed !!!    | " + paddedOutput + " | (Did not copy folder since destination already exists) | " + dest_path)
                
            # However, we should be passing in-non existing directory destnations, since copytree() will create it.
            else:
                logPrint("Copying dir   | " + paddedOutput)
                if (not DEBUG_ONLY_DONT_COPY):
                    shutil.copytree(src_path, dest_path)
    else:
        logPrint("Skipping path | " + paddedOutput + " | (Does not exist)")

def backupRoot():
    logPrint("[ BACKING UP ROOT FOLDERS (Non-user specific) ]")
    
    enableWindowsSleepPrevention(True)

    # Back up selected folders on the root
    displayHeader("Backing up select root folders")

    for path in rootFolderPathsToCopy:
        src_path = path
        dest_path = backupFolderPath + path

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
                logPrint("Skipping path | " + paddedOutput + " | (Is a file, not a folder)")
                continue
            # If the folder is standard
            if filename in standardItemsInRootFolder:
                logPrint("Skipping path | " + paddedOutput + " | (Standard folder)")
                continue
            # If the folder starts with . or $
            if filename.startswith(".") or filename.startswith("$"):
                logPrint("Skipping path | " + paddedOutput + " | (Skipped due to prefix)")
                continue
            # If the folder is hidden
            if(isHidden(fullFilePath)):
                logPrint("Skipping path | " + paddedOutput + " | (Hidden folder)")
                continue
            
            safeCopy(fullFilePath, os.path.join(dest_path, filename))

    enableWindowsSleepPrevention(False)

    logPrint("\n[DONE]")

def backupUser(username):
    logPrint("[ BACKING UP USER: '" + username + "' ]")

    enableWindowsSleepPrevention(True)

    userHomeFolder = getHomeFolderForUsername(username)
    backupUsersFolderPath = os.path.join(backupFolderPath, "Users")

    # Back up each listed folder path
    displayHeader("Backing up user folders and data")
    for relativePath in userFolderPathsToCopy:
        src_path = os.path.join(userHomeFolder, relativePath)
        dest_path = os.path.join(backupUsersFolderPath, username, relativePath)

        safeCopy(src_path, dest_path)
    
    # Back up Work Folders
    # This goes file-by-file, because doing it all at once resulted in errors that were hard to catch and left the backup state unclear.
    workFoldersPath  = os.path.join(userHomeFolder, "Documents", "repos")
    if os.path.exists(workFoldersPath):
        displayHeader("Backing up Work Folders")
        # Finds every single file and directory + sub directories in the entire directory
        for root, dirs, files in os.walk(workFoldersPath, topdown=False):
                # Back up all files
                for name in files:
                    src_path = os.path.join(root, name)
                    dest_path = backupFolderPath + os.path.dirname(os.path.join(root, name))
                    try:
                        safeCopy(src_path, dest_path)
                        Exception()
                    except:
                        logPrint("FAILED COPY!! | " + paddedOutput + " | (Attempted to copy but there was an error)")

                # Back up all directories (that weren't created by the previous files!!!)
                # This should basically create any empty folders
                # This is because any non-empty folders would have been created in the process of copying the files within
                for name in dirs:
                    src_path = os.path.join(root, name)
                    dest_path = backupFolderPath + os.path.join(root, name)
                    # Only back it up if it doesn't exist
                    if(not os.path.exists(dest_path)):
                        safeCopy(src_path, dest_path)

    # Back up .PST files
    src_path = os.path.join(userHomeFolder, "AppData", "Local", "Microsoft", "Outlook")
    dest_path = os.path.join(backupUsersFolderPath, username, "AppData", "Local", "Microsoft", "Outlook")

    if os.path.exists(src_path):
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

    # Back up Firefox/profiles.ini file
    src_path = os.path.join(userHomeFolder, "AppData", "Roaming", "Mozilla", "Firefox", "profiles.ini")
    dest_path = os.path.join(backupUsersFolderPath, username, "AppData", "Roaming", "Mozilla", "Firefox")

    displayHeader("Backing up Mozilla Firefox profiles.ini file")
    safeCopy(src_path, dest_path)

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
                logPrint("Skipping path | " + paddedOutput + " | (Is a file, not a folder)")
                continue
            # If the folder is standard
            if filename in standardItemsInUserFolder:
                logPrint("Skipping path | " + paddedOutput + " | (Standard folder)")
                continue
            # If the folder starts with . or $
            if filename.startswith(".") or filename.startswith("$"):
                logPrint("Skipping path | " + paddedOutput + " | (Skipped due to prefix)")
                continue
            # If the folder is hidden
            if(isHidden(fullFilePath)):
                logPrint("Skipping path | " + paddedOutput + " | (Hidden folder)")
                continue

            safeCopy(fullFilePath, os.path.join(dest_path, filename))

    # Hide AppData folder
    appdataPath = os.path.join(backupUsersFolderPath, username, "AppData")
    displayHeader("Marking AppData folder as hidden " + appdataPath)
    if os.path.exists(appdataPath):
        # Mark as hidden
        ctypes.windll.kernel32.SetFileAttributesW(appdataPath, 2)
        logPrint("AppData folder for '" + username + "' successfully hidden.")
    else:
        logPrint("Error! Could not hide AppData folder-- AppData folder for '" + username + "' does not exist.")
    
    enableWindowsSleepPrevention(False)

    logPrint("\n[DONE]")    

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

