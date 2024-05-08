import tkinter as tk
from tkinter import filedialog
import win32com.client as wc
import os
try:
    inv=wc.GetActiveObject("Inventor.Application")
except:
    inv = wc.Dispatch("Inventor.Application")

inv.Visible = True

def open_ipt_asm_files(folder_path):
 
    for file_name in os.listdir(folder_path):
        if file_name.lower().endswith(('.ipt')):
            file_path = os.path.join(folder_path, file_name)

            try:
                document = inv.Documents.Open(file_path)
                print(f"Opened: {file_path}")
                oContext = inv.TransientObjects.CreateTranslationContext()
                oContext.Type = 13059
                print(oContext)

                oData = inv.TransientObjects.CreateDataMedium()
                directory_path = os.path.join("G:", "021 Python", "HUBBELL", "step")
                file_name = document.DisplayName + ".stp"

                file_path = os.path.join(directory_path, file_name)
                oData.FileName = file_path
                print(oData)

                adin = inv.ApplicationAddIns.ItemById("{90AF7F40-0C01-11D5-8E83-0010B541CD80}") #TranslatorAddIn is class SaveCopyAs is method
                print(adin)

                oOptions = inv.TransientObjects.CreateNameValueMap()
                oOptions.Add("FileFormat", "STEP") 
                '''The Add method is used to add key-value pairs to the NameValueMap. The corrected line now adds the "FileFormat" key with the value "STEP" to the map.oOptions["FileFormat"]="STEP"'''


                print(dir(oOptions))
                ostp = adin.SaveCopyAs(document, oContext, oOptions, oData)
                document.Close(False)
            except Exception as e:
                print(f"Failed to open {file_path}: {str(e)}")

# Replace 'path_to_your_folder' with the actual path to your folder
folder_path = r'G:\003 Inventor\INVENTOR\parts'
open_ipt_asm_files(folder_path)





