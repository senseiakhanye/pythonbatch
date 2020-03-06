import pandas as pd
import json
import shutil
from tkinter import filedialog
from tkinter import *
from pandas import ExcelFile


# top = Tk()
# Code to add widgets will go here...

# var = tk.Label(top, padx = 10, text = "Batch processing for idea")
# var.pack()

# top.filename = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("Excel files","*.xlsx"),("all files","*.*")))
# print(top.filename)
df = pd.read_excel("C:/Users/User/Desktop/testreview.xlsx", sheet_name="Sheet1")
# print(df.columns)

# print(df["info"])
# print(top.filename)

for i, j in df.iterrows():
    # title = j["title"]
    # info = j["info"]
    # instruction = j["instruction"]
    columns = list(j)
    taskList = {"c2JSONObject": TRUE, "data" : {"title": columns[2], "info": columns[3], "instruction": columns[4], "summary": columns[5], "subject": columns[1], "reviews": []}}
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
    print(taskList)
    print()

# top.mainloop()