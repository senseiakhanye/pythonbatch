import pandas as pd
import json
import shutil
import tkinter as tk
import os
from tkinter import filedialog
from pandas import ExcelFile


root = tk.Tk()
canvas1 = tk.Canvas(root, width = 300, height = 300, bg = 'lightsteelblue')
canvas1.pack()
globalFilename = ""
exportFolder = ""

def generateTrueFalse(columns, exportFolder):
    taskList = {"c2JSONObject": 1, "data" : {"title": columns[3], "info": columns[4], "instruction": columns[5], "summary": columns[6], "subject": columns[2], "questions": []}}
    temp = {}
    for x in range(7, len(columns)):
        if (x + 1) % 2 == 0:
            if (pd.isna(columns[x]) == False):
                if (x < len(columns) - 1):
                    temp = {"question" : columns[x], "answer": round(columns[x + 1])}
                    taskList["data"]["questions"].append(temp)
    currentFolder = os.getcwd()
    with open(currentFolder + "/truefalse/ideaData.json", mode="w") as tempFile:
        tempFile.write(json.dumps(taskList, ensure_ascii=False))
    filename = columns[0]
    exportname = exportFolder + "/" + filename
    print(exportname)
    shutil.copytree(currentFolder + "/truefalse", exportname)

def generatePreview(columns, exportFolder):
    taskList = {"c2JSONObject": 1, "data" : {"title": columns[3], "info": columns[4], "instruction": columns[5], "summary": columns[6], "subject": columns[2], "reviews": []}}
    temp = {}
    for x in range(7, len(columns)):
        if (x  + 1) % 2 == 0:
            if (pd.isna(columns[x]) == False):
                if (x < len(columns) - 1):
                    temp = {"point" : columns[x], "description": columns[x + 1]}
                    taskList["data"]["reviews"].append(temp)
    currentFolder = os.getcwd()
    with open(currentFolder +  "/review/ideaData.json", mode="w") as tempFile:
        tempFile.write(json.dumps(taskList, ensure_ascii=False))
    filename = columns[0]
    exportname = exportFolder + "/" + filename
    print(exportname)
    shutil.copytree(currentFolder + "/review", exportname)

def readFile (fileName, exportFolder):
    global globalFilename
    df = pd.read_excel(fileName, sheet_name="Sheet1")
    for i, j in df.iterrows():
        columns = list(j)
        if (columns[1].lower() == "review"):
            generatePreview(columns, exportFolder)
        elif (columns[1].lower() == "true or false"):
            generateTrueFalse(columns, exportFolder)
        
def getExcel ():
    import_file_path = filedialog.askopenfilename()
    if (import_file_path):
        global globalFilename
        globalFilename = import_file_path

def publish():
    global globalFilename
    global exportFolder
    if (globalFilename and exportFolder):
        readFile(globalFilename, exportFolder)
    else:
        print("You have not selected anything")

def folderToExport():
    export_foler = filedialog.askdirectory()
    if (export_foler):        
        global exportFolder
        exportFolder = export_foler
    
publishBtn = tk.Button(text="Process document", command=publish, bg="red", fg="white", font=('helvitica', 12, 'bold'))
browseButton_Excel = tk.Button(text='Import Excel File', command=getExcel, bg='green', fg='white', font=('helvetica', 12, 'bold'))
exportFolderBtn = tk.Button(text="Output folder", command=folderToExport, bg="gray", fg="white", font=('helvitica', 12, 'bold'))
canvas1.create_window(150, 100, window=browseButton_Excel)
canvas1.create_window(150, 150, window=exportFolderBtn)
canvas1.create_window(150, 200, window=publishBtn)

root.mainloop()