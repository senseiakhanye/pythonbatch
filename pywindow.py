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

def genrateMCQ(columns, exportFolder):
    taskList = {"c2JSONObject": 1, "data" : {"title": columns[3], "startScreen": columns[4], "instruction": columns[5], "subject": columns[2], "questions": []}}
    temp = {}
    for x in range(6, len(columns), 5):
        if (pd.isna(columns[x]) == False):
            if (x < len(columns) - 4):
                temp = {"question" : columns[x].replace('’', '\''), "option1": str(columns[x + 1]), "option2": str(columns[x + 2]), "option3": str(columns[x + 3]), "option4": str(columns[x + 4])}
                taskList["data"]["questions"].append(temp)
    currentFolder = os.getcwd()
    with open(currentFolder + "/mcq/ideadata.json", mode="w") as tempFile:
        tempFile.write(json.dumps(taskList, ensure_ascii=False))
    filename = columns[0]
    exportname = exportFolder + "/" + filename
    shutil.copytree(currentFolder + "/mcq", exportname)

def generateTrueFalse(columns, exportFolder):
    taskList = {"c2JSONObject": 1, "data" : {"title": columns[3], "info": columns[4], "instruction": columns[5], "summary": columns[6], "subject": columns[2], "questions": []}}
    temp = {}
    for x in range(7, len(columns)):
        if (x + 1) % 2 == 0:
            if (pd.isna(columns[x]) == False):
                if (x < len(columns) - 1):
                    temp = {"question" : columns[x].replace('’', '\''), "answer": round(columns[x + 1])}
                    taskList["data"]["questions"].append(temp)
    currentFolder = os.getcwd()
    with open(currentFolder + "/truefalse/ideadata.json", mode="w") as tempFile:
        tempFile.write(json.dumps(taskList, ensure_ascii=False))
    filename = columns[0]
    exportname = exportFolder + "/" + filename
    shutil.copytree(currentFolder + "/truefalse", exportname)

def generatePreview(columns, exportFolder):
    taskList = {"c2JSONObject": 1, "data" : {"title": columns[3], "info": columns[3].replace('’', '\''), "instruction": columns[4].replace('’', '\''), "summary": columns[6].replace('’', '\'').replace('\n\n','`'), "summary-heading": columns[5].replace('’', '\''),"subject": columns[2], "reviews": []}}
    temp = {}
    for x in range(7, len(columns)):
        if (x  + 1) % 2 == 0:
            if (pd.isna(columns[x]) == False and columns[x] and len(columns[x].strip()) > 0):
                if (x < len(columns) - 1):
                    temp = {"point" : columns[x].replace('’', '\''), "description": columns[x + 1].replace('’', '\'').replace('\n\n','`')}
                    taskList["data"]["reviews"].append(temp)
    currentFolder = os.getcwd()
    with open(currentFolder +  "/review/ideadata.json", mode="w") as tempFile:
        tempFile.write(json.dumps(taskList, ensure_ascii=False))
    filename = columns[0]
    exportname = exportFolder + "/" + filename
    shutil.copytree(currentFolder + "/review", exportname)

def generateLearningGoal(columns, exportFolder):
    taskList = {"c2JSONObject": 1, "data" : {"title": columns[3], "instruction": columns[4].replace('’', '\''), "subject": columns[2], "goals": []}}
    temp = {}
    learningGoals = 0
    for x in range(5, len(columns), 3):
        if (pd.isna(columns[x]) == False and columns[x] and len(columns[x].strip()) > 0):
            if (x < len(columns) - 1):
                temp = {"point" : columns[x].replace('’', '\'').replace("\n\n","\n"), "description": columns[x + 2].replace('’', '\'').replace("\n\n","\n"), "heading": columns[x + 1].replace('’', "heading").replace("\n\n","\n")}
                taskList["data"]["goals"].append(temp)
                learningGoals += 1
    currentFolder = os.getcwd()
    filename = columns[0]
    os.makedirs(exportFolder + "/" + filename)
    for lg in range(-1, learningGoals):        
        newList = taskList.copy()
        newList["data"]["active"] = str(lg)
        exportname = exportFolder + "/" + filename + "/" + filename + "_" + str(lg + 1)
        shutil.copytree(currentFolder + "/learninggoal", exportname)
        with open(exportname + "/ideadata.json", mode="w") as tempFile:
            tempFile.write(json.dumps(newList, ensure_ascii=False))
        outputFilename = exportname
        shutil.make_archive(outputFilename, 'zip', exportname + "/")

def readFile (fileName, exportFolder):
    global globalFilename
    df = pd.read_excel(fileName, sheet_name="Review")
    for i, j in df.iterrows():
        columns = list(j)
        if (pd.isna(columns[1]) == False):
            if (columns[1].lower() == "review"):
                generatePreview(columns, exportFolder)
            elif (columns[1].lower() == "true or false"):
                generateTrueFalse(columns, exportFolder)
            elif (columns[1].lower() == "learning goal"):
                generateLearningGoal(columns, exportFolder)
            elif (columns[1].lower() == "mcq"):
                genrateMCQ(columns, exportFolder)

        
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