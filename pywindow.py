import pandas as pd
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
    taskList = {"c2JSONObject": TRUE, "data" : {"title": columns[0], "info": columns[1], "instruction": columns[2], "reviews": []}}
    n = 3
    temp = {}
    for x in range(3, len(columns)):
        if x % 3 == 0:
            temp = {"point" : columns[x], "description": columns[x + 1]}
            taskList["data"]["reviews"].append(temp)
        with open("C:/Users/User/Desktop/ideaData.json", mode="w") as tempFile:
            tempFile.write(str(taskList))
    # print(title + " " + info + " " + instruction)
    print(taskList)
    
    print()

# top.mainloop()