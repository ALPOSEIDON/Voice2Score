import MyPackages as mp
import MyGUI as mg
from PySide6.QtWidgets import QApplication

# mp.file_store(r"Resources\backup\namelists.json", ['那好','hello world'])

# name = mp.file_loader(r"Resources\backup\namelists.json")

# print(name)

# print(name[0])

# excel = mp.find_xlsx_files('Excel')
# print(excel)

# sheets =  mp.find_xlsx_Sheets('Excel\\'+excel[0])
# print(sheets)




dic = {
    "root_directory" : "Excel", 
    "file_name":"chinese.xlsx",
    "whole_file_path" : r"Excel\chinese.xlsx",
    "sheet_name" : "Sheet1",
    "NAME_Init_mode" : 0,
    "NAME_COL" : "B",
    "NAME_ROW_START" : 3,
    "NAME_ROW_END" : 57,
    "init_mode" : 0,
    "usingAudioRecognition" : False,
    "ColScore" : ""
}

# mp.file_store(r'Resources\backup\settings.json', dic)

# mp.refresh_namelists("B", 3, 57, r"Excel\chinese.xlsx", "Sheet1", r"Resources\backup\namelists.json")

app = QApplication()
window = mg.MWindow()

# window.ui.nameBuffer.addItem("当前学生人数："+str(len(name)))
# window.ui.nameBuffer.addItems(name)
#for na in name:
#    window.ui.nameBuffer.addItem(str(na))

# window.ui.studentScore.display(0.0)

# print(mp.cmd_analyse_name_score("你好123"))

window.show()
window.statusTip()

app.exec()