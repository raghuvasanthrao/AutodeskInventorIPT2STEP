import tkinter as tk
from tkinter import filedialog
import win32com.client as wc
import os

try:
    inv = wc.GetActiveObject("Inventor.Application")
except:
    inv = wc.Dispatch("Inventor.Application")

inv.Visible = True

def open_ipt_asm_files(source_folder_path, destination_directory_path):
    for file_name in os.listdir(source_folder_path):
        if file_name.lower().endswith('.ipt'):
            file_path = os.path.join(source_folder_path, file_name)

            try:
                document = inv.Documents.Open(file_path)
                print(f"Opened: {file_path}")

                oContext = inv.TransientObjects.CreateTranslationContext()
                oContext.Type = 13059
                print(oContext)

                oData = inv.TransientObjects.CreateDataMedium()
                file_name_stp = document.DisplayName + ".stp"
                file_path_stp = os.path.join(destination_directory_path, file_name_stp)
                oData.FileName = file_path_stp
                print(oData)

                adin = inv.ApplicationAddIns.ItemById("{90AF7F40-0C01-11D5-8E83-0010B541CD80}")
                print(adin)

                oOptions = inv.TransientObjects.CreateNameValueMap()
                oOptions.Add("FileFormat", "STEP")

                print(dir(oOptions))
                ostp = adin.SaveCopyAs(document, oContext, oOptions, oData)
                document.Close(False)

            except Exception as e:
                print(f"Failed to open {file_path}: {str(e)}")

def browse_source_folder():
    folder_path = filedialog.askdirectory(title="Select Source Folder")
    source_folder_var.set(folder_path)

def browse_destination_directory():
    directory_path = filedialog.askdirectory(title="Select Destination Directory")
    destination_directory_var.set(directory_path)

# Create the main window
root = tk.Tk()
root.title("Inventor Script")

# Variables to store folder paths
source_folder_var = tk.StringVar()
destination_directory_var = tk.StringVar()

# Entry widgets to display the selected folder paths
source_folder_entry = tk.Entry(root, textvariable=source_folder_var, width=50)
destination_directory_entry = tk.Entry(root, textvariable=destination_directory_var, width=50)

# Buttons to open file dialogs
source_folder_button = tk.Button(root, text="Browse Source Folder", command=browse_source_folder)
destination_directory_button = tk.Button(root, text="Browse Destination Directory", command=browse_destination_directory)

# Button to execute the script
execute_button = tk.Button(root, text="Execute Script", command=lambda: open_ipt_asm_files(source_folder_var.get(), destination_directory_var.get()))

# Label at the bottom
developed_by_label = tk.Label(root, text="Developed by Raghu Vasanth Rao", font=("Arial", 8), fg="white")

# Pack widgets into the window
source_folder_entry.pack(pady=5)
source_folder_button.pack(pady=5)
destination_directory_entry.pack(pady=5)
destination_directory_button.pack(pady=5)
execute_button.pack(pady=10)
developed_by_label.pack(side=tk.BOTTOM, pady=10)

# Run the Tkinter event loop
root.mainloop()
