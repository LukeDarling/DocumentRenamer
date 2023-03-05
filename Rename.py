# Written by Luke Darling

# Imports
from tkinter import *
from tkinter import filedialog, messagebox
import os, shutil

# Define constants
IMPORT_FOLDER_NAME = "Raw"
EXPORT_FOLDER_NAME = "Done"

# Request folder to operate on
folder = filedialog.askdirectory()

# Check whether selected folder is the root folder
# If not, move up a directory
(parentFolder, folderName) = os.path.split(folder)
if folderName.lower() == IMPORT_FOLDER_NAME.lower() or folderName.lower() == EXPORT_FOLDER_NAME.lower():
    (parentFolder, folderName) = os.path.split(parentFolder)

# Define project attributes
projectName = folderName
projectPath = os.path.join(parentFolder, folderName)
projectImportPath = os.path.join(projectPath, IMPORT_FOLDER_NAME)
projectExportPath = os.path.join(projectPath, EXPORT_FOLDER_NAME)

# Get filenames from import directory
filenames = os.listdir(projectImportPath)

# Define file lists
renameList = []
skipList = []
failureList = []

# Generate new filenames and skip existing 
for filename in filenames:
    file = {"oldName": filename, "oldPath": os.path.join(projectImportPath, filename)}
    file["newName"] = filename.split("_")[0] + "_" + projectName + os.path.splitext(file["oldName"])[1]
    file["newPath"] = os.path.join(projectExportPath, file["newName"])
    if os.path.exists(file["newPath"]):
        skipList.append(file)
    else:
        renameList.append(file)

for file in renameList:
    try:
        shutil.copy2(file["oldPath"], file["newPath"])
    except:
        failureList.append(file)

for file in failureList:
    renameList.remove(file)

message = ""
if len(renameList) > 0:
    message += str(len(renameList)) + " document" + ("" if len(renameList) == 1 else "s") + " successfully copied and renamed."
if len(skipList) > 0:
    if message != "":
        message += "\n\n"
    message += str(len(skipList)) + " document" + ("" if len(skipList) == 1 else "s") + " skipped."
if len(failureList) > 0:
    if message != "":
        message += "\n\n"
        message += str(len(failureList)) + " document" + ("" if len(failureList) == 1 else "s") + " failed."
if message == "":
    message = "No documents found."

if len(failureList) > 0:
    messagebox.showerror("Rename", message)
elif len(skipList) > 0:
    messagebox.showwarning("Rename", message)
elif len(renameList) > 0:
    messagebox.showinfo("Rename", message)
else:
    messagebox.showwarning("Rename", message)
