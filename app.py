from win32com.shell import shell, shellcon
import os
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import macro
import shutil



print shell.SHGetFolderPath(0, shellcon.CSIDL_DESKTOP, None, 0)
os.system("open -a 'path/FolderPath Excel.app' 'path/learning goals.xlsx'")
df = pd.read_excel('learning goals.xlsx', sheetname='Sheet1')

folder_obj = learning_goals.cell(row=1, column=1)
distribution_obj = learning_goals.cell(row=1, column=2)
db = learning_goals.cell(columnletter + str(wb.sheets[sheetname].cells.last_cell.row)

class Data(db.Model):
    __tablename__="data"
    id=db.Column(db.Integer, primary_key=True)
    key_=db.Column(db.String(120), unique=True)
    height_=db.Column(db.Integer)

    def __init__(self, key, height_):
        self.key_=key_
        self.height_=height_

file_path = "/my/directory/filename.txt"
directory = os.path.dirname(file_path)

try:
    os.stat(directory)
except:
    os.mkdir(directory)       

f = file(folder_obj)

def index():
    return render_template("index.html")


def success():
    if request.method=='distribution_obj':
        key=request.form["key_name"]
        height=request.form["height_name"]
        print(key, height)
        if db.session.query(Data).filter(Data.key_ == key).count()== 0:
            data=Data(new.js,height)
            session.add(data)
            session.commit()
            average_height=db.session(func.avg(Data.height_)).scalar()
            average_height=round(average_height, 1
            print(average_height)
            return render_template("success.html")
    return render_template('index.html')

if __name__ == '__main__':
    app.debug=True
    newPath = shutil.copy('data.js', '/desktop/automate')
    
