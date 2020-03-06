import pandas as pd
import json
import shutil
import tkinter as tk
from tkinter import filedialog
from pandas import ExcelFile


root = tk.Tk()
canvas1 = tk.Canvas(root, width = 300, height = 300, bg = 'lightsteelblue')
canvas1.pack()
globalFilename = ""

def readFile (fileName):
    df = pd.read_excel(fileName, sheet_name="Sheet1")
    for i, j in df.iterrows():
        columns = list(j)
        taskList = {"c2JSONObject": 1, "data" : {"title": columns[2], "info": columns[3], "instruction": columns[4], "summary": columns[5], "subject": columns[1], "reviews": []}}
        temp = {}
        for x in range(6, len(columns)):
            if x % 2 == 0:
                if (x < len(columns) - 1):
                    temp = {"point" : columns[x], "description": columns[x + 1]}
                    taskList["data"]["reviews"].append(temp)
        with open("C:/Users/User/Desktop/template/ideaData.json", mode="w") as tempFile:
            tempFile.write(json.dumps(taskList, ensure_ascii=False))
        filename = columns[0]
        shutil.copytree("c:/Users/User/Desktop/template", "c:/Users/User/Desktop/templates/" + filename)
    # root.destroy()

def getExcel ():
    import_file_path = filedialog.askopenfilename()
    if (import_file_path):
        readFile(import_file_path)
    
publishBtn = tk.Button(text="Process document", command=readFile, bg="red", fg="white", font=('helvitica', 12, 'bold'))
browseButton_Excel = tk.Button(text='Import Excel File', command=getExcel, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 150, window=browseButton_Excel)
canvas1.create_window(150, 200, window=publishBtn)

root.mainloop()