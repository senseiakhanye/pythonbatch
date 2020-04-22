import pandas as pd
import json
import shutil
import tkinter as tk
import os
import sys
import copy
import datetime
from tkinter import filedialog
from pandas import ExcelFile
from tkinter import ttk
from tkinter import messagebox
from platform import system
from pathlib import Path



root = tk.Tk()
canvas1 = tk.Canvas(root, width = 300, height = 300, bg = 'lightsteelblue')

canvas1.pack()
globalFilename = ""
exportFolder = ""
showProgress = True

def getSafeString(st):
    return str(st).replace('’', '\'').replace("\n\n","\n").replace("–","-").replace("‘", "'")

def getCorrrectFilename(folder, filename):
    x = 0
    tempfilename = filename
    while (os.path.isdir(os.path.join(folder, tempfilename)) == True):
        x += 1
        tempfilename = filename + "_" + str(x)
    return tempfilename

def combineOption(corrVal, textVal):
    if (isinstance(textVal, datetime.datetime)):
        textVal = textVal.strftime('%d %B %Y')
    return getSafeString(str(corrVal)) + "|" + getSafeString(str(textVal))

def getDataDirectory():
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.getcwd()
    return base_path

def mcqNotEmpty(columns):
    return (pd.isna(columns[1]) == False and pd.isna(columns[3]) == False and pd.isna(columns[5]) == False and pd.isna(columns[7]) == False and pd.isna(columns[9]) == False)

def reportNotEmpty(columns, x):
    return (pd.isna(columns[x]) == False and pd.isna(columns[x + 1]) == False and pd.isna(columns[x + 2]) == False)

def multipleChoiceNotEmpty(columns, x):
    return (pd.isna(columns[x]) == False and pd.isna(columns[x + 1]) == False)

def generateMCQAssessments(df, exportFolder):
    taskList = {"c2JSONObject": 1, "data" : {}}
    temp = {}
    currentFolder = getDataDirectory()
    allquestion = []
    filename = getCorrrectFilename(exportFolder, "testfilename")
    for i, j in df.iterrows():
        if (i == 1):
            columns = list(j)
            filename = getCorrrectFilename(exportFolder, columns[0])
            taskList["data"] = {"title": columns[6], "instruction": columns[8], "subject": columns[4]}
        elif (i >= 3):
            columns = list(j)
            if  mcqNotEmpty(columns):
                temp = {"question" : str(columns[1]).replace('’', '\''), "option1": combineOption(columns[2], columns[3]), "option2": combineOption(columns[4], columns[5]), "option3":combineOption(columns[6], columns[7]), "option4": combineOption(columns[8], columns[9])}
                tempDictionary = {}
                tempDictionary.update(taskList)
                tempDictionary["data"].update(temp)
                allquestion.append(copy.deepcopy(tempDictionary))
    os.makedirs(exportFolder + "/" + filename)
    for i, obj in enumerate(allquestion):
        exportname = exportFolder + "/" + filename + "/" + filename + "_" + str(i)
        shutil.copytree(currentFolder + "/plugins/mcq-assessment", exportname)
        with open(exportname + "/ideadata.json", mode="w") as tempFile:
            tempFile.write(json.dumps(obj, ensure_ascii=False))
        outputFilename = exportname
        shutil.make_archive(outputFilename, 'zip', exportname + "/")

def generateMCQ(df, exportFolder):
    taskList = {"c2JSONObject": 1, "data" : {}}
    temp = {}
    currentFolder = getDataDirectory()
    filename = getCorrrectFilename(exportFolder, "testfilename")
    for i, j in df.iterrows():
        if (i == 1):
            columns = list(j)
            filename = getCorrrectFilename(exportFolder, columns[0])
            taskList["data"] = {"title": columns[6], "startScreen": columns[10], "instruction": columns[8], "subject": columns[4], "questions": []}
        elif (i >= 3):
            columns = list(j)
            if  mcqNotEmpty(columns):
                temp = {"question" : getSafeString(columns[1]), "option1": combineOption(columns[2], columns[3]), "option2": combineOption(columns[4], columns[5]), "option3":combineOption(columns[6], columns[7]), "option4": combineOption(columns[8], columns[9])}
                taskList["data"]["questions"].append(temp)
    with open(currentFolder + "/plugins/mcq/ideadata.json", mode="w") as tempFile:
        tempFile.write(json.dumps(taskList, ensure_ascii=False))
    exportname = exportFolder + "/" + filename
    shutil.copytree(currentFolder + "/plugins/mcq", exportname)
    outputFilename = exportname
    shutil.make_archive(outputFilename, 'zip', exportname + "/")

def genrateAllMCQ(df, exportFolder):
    global showProgress
    filename = getCorrrectFilename(exportFolder , "MCQ")
    fullFolder = exportFolder + "/" + filename
    os.makedirs(fullFolder)
    os.makedirs(fullFolder + "/Assessments")
    os.makedirs(fullFolder + "/Interactives")
    generateMCQAssessments(df, fullFolder + "/Assessments")
    generateMCQ(df, fullFolder + "/Interactives")
    if (showProgress):
        messagebox.showinfo("Generated", "Generated MCQs: (" + os.path.join(exportFolder, filename) + ")")

def generateMCInteractive(columns, exportFolder):
    taskList = {"c2JSONObject": 1, "data" : {"title": getSafeString(columns[3]), "startScreen": getSafeString(columns[4]), "instruction": getSafeString(columns[5]), "subject": columns[2], "points": []}}
    temp = {}
    cnt = 0
    maxNum = 8
    for x in range(6, len(columns), 2):
        if multipleChoiceNotEmpty(columns, x) and cnt < maxNum:
            temp = {"question" : getSafeString(columns[x]), "answer": getSafeString(columns[x + 1])}
            taskList["data"]["points"].append(temp)
            cnt += 1
    currentFolder = getDataDirectory()
    with open(currentFolder +  "/plugins/match/ideadata.json", mode="w") as tempFile:
        tempFile.write(json.dumps(taskList, ensure_ascii=False))
    # filename = columns[0]
    filename = getCorrrectFilename(exportFolder, columns[0])
    exportname = exportFolder + "/" + filename
    shutil.copytree(currentFolder + "/plugins/match", exportname)
    outputFilename = exportname
    shutil.make_archive(outputFilename, 'zip', exportname + "/")
    # messagebox.showinfo("Generated", "Generated Learning goals : (" + os.path.join(exportFolder, filename) + ")")

def generateMatchingColumns(df, exportFolder, func):
    for i, j in df.iterrows():
        columns = list(j)
        if (pd.isna(columns[1]) == False):
            if (columns[1].lower() == "matching columns"):
                func(columns, exportFolder)

def genrateAllMatchingColumns(df, exportFolder):
    global showProgress
    filename = getCorrrectFilename(exportFolder , "Matching Columns")
    fullFolder = exportFolder + "/" + filename
    os.makedirs(fullFolder)
    os.makedirs(fullFolder + "/Assessments")
    os.makedirs(fullFolder + "/Interactives")
    # generateMCInteractive(df, fullFolder + "/Assessments")
    generateMatchingColumns(df, fullFolder + "/Interactives", generateMCInteractive)
    # generateMCQ(df, fullFolder + "/Interactives")
    if (showProgress):
        messagebox.showinfo("Generated", "Generated Multiple choice quesitons: (" + os.path.join(exportFolder, filename) + ")")

def generateComprehension(columns, exportFolder, assessmentType):
    print(columns[5])
    taskList = {"c2JSONObject": 1, "data" : {"title": getSafeString(columns[3]), "startScreen": "", "instruction": getSafeString(columns[4]), "subject": columns[2], "summary": getSafeString(columns[5]), "question": getSafeString(columns[6]), "assessment": str(assessmentType),"options": []}}
    temp = {}
    cnt = 0
    for x in range(7, len(columns)):
        if pd.isna(columns[x]) == False and len(getSafeString(columns[x]).strip()) > 0 :
            isTrue = 0
            if (cnt == 0):
                isTrue = 1
            temp = {"option" : combineOption(str(isTrue), getSafeString(columns[x]))}
            taskList["data"]["options"].append(temp)
            cnt += 1
    currentFolder = getDataDirectory()
    with open(currentFolder +  "/plugins/comprehension/ideadata.json", mode="w") as tempFile:
        tempFile.write(json.dumps(taskList, ensure_ascii=False))
    # filename = columns[0]
    filename = getCorrrectFilename(exportFolder, columns[0])
    exportname = exportFolder + "/" + filename
    shutil.copytree(currentFolder + "/plugins/comprehension", exportname)
    outputFilename = exportname
    shutil.make_archive(outputFilename, 'zip', exportname + "/")
    # messagebox.showinfo("Generated", "Generated Learning goals : (" + os.path.join(exportFolder, filename) + ")")

def generateComprehensions(df, exportFolder, assessmentType):
    for i, j in df.iterrows():
        columns = list(j)
        if (pd.isna(columns[1]) == False):
            if (columns[1].lower() == "comprehension"):
                generateComprehension(columns, exportFolder, assessmentType)

def genrateAllComprehensions(df, exportFolder):
    global showProgress
    filename = getCorrrectFilename(exportFolder , "Comprehensions")
    fullFolder = exportFolder + "/" + filename
    os.makedirs(fullFolder)
    os.makedirs(fullFolder + "/Assessments")
    os.makedirs(fullFolder + "/Interactives")
    generateComprehensions(df, fullFolder + "/Interactives", 0)
    generateComprehensions(df, fullFolder + "/Assessments", 1)
    if (showProgress):
        messagebox.showinfo("Generated", "Generated Multiple choice quesitons: (" + os.path.join(exportFolder, filename) + ")")

def generateTrueFalse(columns, exportFolder):
    global showProgress
    taskList = {"c2JSONObject": 1, "data" : {"title": columns[3], "info": columns[4], "instruction": columns[5], "summary": columns[6], "subject": columns[2], "questions": []}}
    temp = {}
    for x in range(7, len(columns)):
        if (x + 1) % 2 == 0:
            if (pd.isna(columns[x]) == False):
                if (x < len(columns) - 1):
                    temp = {"question" : columns[x].replace('’', '\''), "answer": round(columns[x + 1])}
                    taskList["data"]["questions"].append(temp)
    currentFolder = getDataDirectory()
    with open(currentFolder + "/truefalse/ideadata.json", mode="w") as tempFile:
        tempFile.write(json.dumps(taskList, ensure_ascii=False))
    # filename = columns[0]
    filename = getCorrrectFilename(exportFolder, columns[0])
    exportname = exportFolder + "/" + filename
    shutil.copytree(currentFolder + "/truefalse", exportname)
    if (showProgress):
        messagebox.showinfo("Generated", "Generated True or False assets : (" + os.path.join(exportFolder, filename) + ")")

def generatePreview(columns, exportFolder):
    global showProgress
    taskList = {"c2JSONObject": 1, "data" : {"title": columns[3], "info": columns[3].replace('’', '\''), "instruction": columns[4].replace('’', '\''), "summary": columns[6].replace('’', '\'').replace('\n\n','`'), "summary-heading": columns[5].replace('’', '\''),"subject": columns[2], "reviews": []}}
    temp = {}
    for x in range(7, len(columns)):
        if (x  + 1) % 2 == 0:
            if (pd.isna(columns[x]) == False and columns[x] and len(columns[x].strip()) > 0):
                if (x < len(columns) - 1):
                    temp = {"point" : columns[x].replace('’', '\''), "description": columns[x + 1].replace('’', '\'').replace('\n\n','`')}
                    taskList["data"]["reviews"].append(temp)
    currentFolder = getDataDirectory()
    with open(currentFolder +  "/review/ideadata.json", mode="w") as tempFile:
        tempFile.write(json.dumps(taskList, ensure_ascii=False))
    # filename = columns[0]
    filename = getCorrrectFilename(exportFolder, columns[0])
    exportname = exportFolder + "/" + filename
    shutil.copytree(currentFolder + "/review", exportname)
    if (showProgress):
        messagebox.showinfo("Generated", "Generated reviews : (" + os.path.join(exportFolder, filename) + ")")

def generateLearningGoal(columns, exportFolder):
    global showProgress
    taskList = {"c2JSONObject": 1, "data" : {"title": getSafeString(columns[3]), "instruction": getSafeString(columns[4].replace('’', '\'')), "subject": getSafeString(columns[2]), "goals": []}}
    temp = {}
    learningGoals = 0
    for x in range(5, len(columns), 2):
        if (pd.isna(columns[x]) == False and columns[x] and len(columns[x].strip()) > 0):
            temp = {"point" : getSafeString(columns[x]), "description": getSafeString(columns[x + 1]), "heading": "After completing this subtopic you will be able to:"}
            taskList["data"]["goals"].append(temp)
            learningGoals += 1
    currentFolder = getDataDirectory()
    # filename = columns[0]
    filename = getCorrrectFilename(exportFolder, columns[0])
    os.makedirs(exportFolder + "/" + filename)
    for lg in range(-1, learningGoals):        
        newList = taskList.copy()
        newList["data"]["active"] = str(lg)
        exportname = exportFolder + "/" + filename + "/" + filename + "_" + str(lg + 1)
        shutil.copytree(currentFolder + "/plugins/learninggoal", exportname)
        with open(exportname + "/ideadata.json", mode="w") as tempFile:
            tempFile.write(json.dumps(newList, ensure_ascii=False))
        outputFilename = exportname
        shutil.make_archive(outputFilename, 'zip', exportname + "/")
    if (showProgress):
        messagebox.showinfo("Generated", "Generated Learning goals : (" + os.path.join(exportFolder, filename) + ")")

def generateReports(columns, exportFolder):
    global showProgress
    taskList = {"c2JSONObject": 1, "data" : {"title": columns[3], "instruction": getSafeString(columns[4]), "subject": columns[2], "reviews": []}}
    temp = {}
    for x in range(5, len(columns), 3):
        if reportNotEmpty(columns, x):
            temp = {"point" : getSafeString(columns[x]), "heading": getSafeString(columns[x + 1]), "description": getSafeString(columns[x + 2])}
            taskList["data"]["reviews"].append(temp)
    currentFolder = getDataDirectory()
    with open(currentFolder +  "/plugins/report/ideadata.json", mode="w") as tempFile:
        tempFile.write(json.dumps(taskList, ensure_ascii=False))
    # filename = columns[0]
    filename = getCorrrectFilename(exportFolder, columns[0])
    exportname = exportFolder + "/" + filename
    shutil.copytree(currentFolder + "/plugins/report", exportname)
    outputFilename = exportname
    shutil.make_archive(outputFilename, 'zip', exportname + "/")
    if (showProgress):
        messagebox.showinfo("Generated", "Generated index and info : (" + os.path.join(exportFolder, filename) + ")")
    # messagebox.showinfo("Generated", "Generated Learning goals : (" + os.path.join(exportFolder, filename) + ")")

def generateReportsUpdated(df, exportFolder):
    filename = getCorrrectFilename(exportFolder , "Index and info")
    fullFolder = exportFolder + "/" + filename
    os.makedirs(fullFolder)
    for i, j in df.iterrows():
        columns = list(j)
        if (pd.isna(columns[1]) == False):
            if (columns[1].lower() == "index and info" or columns[1].lower() == "index & info"):
                generateReports(columns, fullFolder)

def generateLearningGoalUpdated(df, exportFolder):
    filename = getCorrrectFilename(exportFolder , "Learning goals")
    fullFolder = exportFolder + "/" + filename
    os.makedirs(fullFolder)
    for i, j in df.iterrows():
        columns = list(j)
        if (pd.isna(columns[1]) == False):
            if (columns[1].lower() == "learning goal" or columns[1].lower() == "learning goals"):
                generateLearningGoal(columns, fullFolder)

def readFile (fileName, sheetName, exportFolder):
    global globalFilename    
    df = 0
    try:
        df = pd.read_excel(fileName, sheetName)
    except Exception as e:
        print(e)
        messagebox.showerror("Error", "Please rename the sheet to process: (" + sheetName + ")")
        return
    # if sheetName.lower() == "mcq assessment":
    #     genrateMCQ(df, exportFolder)
    # else:
    firstRowData = df.head(0)
    headings = list(firstRowData)
    if (headings[0].lower().strip() == "mcq assessment" or headings[0].lower().strip() == "mcq"):
        genrateAllMCQ(df, exportFolder)
        return
    elif (headings[0].lower() == "learning goals"):
        generateLearningGoalUpdated(df, exportFolder)
        return
    elif (headings[0].lower().strip() == "index and info"):
        generateReportsUpdated(df, exportFolder)
        return
    elif (headings[0].lower().strip() == "matching columns"):
        genrateAllMatchingColumns(df, exportFolder)
        return
    elif (headings[0].lower().strip() == "comprehension"):
        genrateAllComprehensions(df, exportFolder)
        return
    for i, j in df.iterrows():
        columns = list(j)
        if (pd.isna(columns[1]) == False):
            if (columns[1].lower() == "review"):
                generatePreview(columns, exportFolder)
            elif (columns[1].lower() == "true or false"):
                generateTrueFalse(columns, exportFolder)
            elif (columns[1].lower() == "learning goal"):
                generateLearningGoal(columns, exportFolder)
            # elif (columns[1].lower() == "mcq assessment"):
            #     genrateMCQ(columns, exportFolder)

        
def getExcel ():
    import_file_path = filedialog.askopenfilename()
    if (import_file_path):
        global globalFilename
        # global comboExample
        globalFilename = import_file_path
        df = pd.read_excel(globalFilename, None)
        myList = ["All"]
        
        for k in df.keys():
            myList.append(k)
        comboExample["values"] = myList
        comboExample.current(0)
        canvas1.create_window(150, 100, window=comboExample)

def generateAll(filename, exportFolder):
    global showProgress
    showProgress = False
    df = pd.read_excel(filename, None)
    for k in df.keys():
        print("Generating : " + k)
        readFile(filename, k, exportFolder)
    messagebox.showinfo("Generated", "All Spreadsheet processed : (" + exportFolder + ")")

def publish():
    global globalFilename
    global exportFolder
    if (globalFilename == ""):
        messagebox.showerror("Error", "Please select excel file to process")
        return
    if (exportFolder == ""):
        messagebox.showerror("Error", "Please select export folder")
        return

    comboBoxValue = comboExample.get()
    sheetName = comboBoxValue
    parentFolderName = Path(globalFilename).stem
    filename = getCorrrectFilename(exportFolder, parentFolderName)
    exportFolder = exportFolder + "/" + filename
    os.makedirs(exportFolder)
    if (comboExample.get() == "All"):
        generateAll(globalFilename, exportFolder)
    else:
        readFile(globalFilename, sheetName, exportFolder)

def folderToExport():
    export_foler = filedialog.askdirectory()
    if (export_foler):
        global exportFolder
        exportFolder = export_foler

ttk.Style().configure('green/black.TLabel', foreground='green', background='black')
ttk.Style().configure('green/black.TButton', foreground='blue', background='black')
comboExample = ttk.Combobox(root, values=[], state='readonly')
publishBtn = ttk.Button(text="Process document", command=publish, style="green/black.TButton")
browseButton_Excel = ttk.Button(text='Import Excel File', command=getExcel, style="green/black.TButton")
exportFolderBtn = ttk.Button(text="Output folder", command=folderToExport, style="green/black.TButton")
canvas1.create_window(150, 50, window=browseButton_Excel)
canvas1.create_window(150, 150, window=exportFolderBtn)
canvas1.create_window(150, 200, window=publishBtn)

root.title('Idea batch asset creator')
platformD = system()
if platformD == 'Windows':
    iconLogo = getDataDirectory() + "/PluginIcon.ico"
    root.iconbitmap(iconLogo)
else:
    iconLogo = getDataDirectory() + "/PluginIcon.png"
    # img = tk.Image("photo", file= iconLogo)
    # root.iconphoto(True, img) # you may also want to try this.
    # root.tk.call('wm','iconphoto', root._w, img)
    # print(iconLogo)

root.mainloop()