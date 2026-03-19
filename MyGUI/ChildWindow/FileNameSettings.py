from PySide6.QtWidgets import QMessageBox, QFileDialog, QDialog, QApplication
from PySide6.QtUiTools import QUiLoader
import re
import MyPackages as mp

class FileNameSettingsWindow(QDialog):

    returnValueTuple = (
            "FileName",
            "SheetName",
            "ColName",
            "RowStart",
            "RowEnd",
            "ColScore",
        )

    def __init__(self, settings:dict):
        super().__init__()
        try:
            # 返回的对象是一个 QMainWindow类的实例 也就是我创建的主窗口
            self.ui = QUiLoader().load(r'Resources\backup\FileNameSetting.ui')
        except Exception as e:
            QMessageBox.critical(None, "初始化异常", 
                                 f"读取失败：{type(e).__name__}\n{e}\n请检查 Resources\\backup\FileNameSetting.ui 文件是否出现问题")
        
        ## 把 UI 的布局加载进去
        lay = self.ui.layout()
        self.setLayout(lay)
        self.resize(480, 462)
        self.move(300, 200)
        self.setWindowTitle("默认设置")

        ## 先把一些参数初始化
        self.settingdict = settings.copy()

        self.isFileOK = False
        self.isSheetOK = False
        self.isColOK = False
        self.isNameOK = False

        self.filePath = self.settingdict["whole_file_path"]                          # 选择的文件路径
        self.SheetName = self.settingdict["sheet_name"]                         # 选择的Sheet名
        self.ColName = ""                           # 选择的Col名
        self.RowStart = 0
        self.RowEnd = 0
        self.ColScore = ""

        ## ---------------------------- 信号绑定 ----------------------------
        self.ui.buttonBox.accepted.connect(self.__close_check)  # 可以连接到我自己的函数上，方便self.accept()自动调用
        self.ui.buttonBox.rejected.connect(self.reject)

        self.ui.fileChoose.clicked.connect(self.__choose_file_dialog)

        ## 对界面进行初始化
        self.__init_all_uis()

    ### ———————————————————————————— 操作和检验函数 ————————————————————————————
    def __choose_file_dialog(self):
        file_path, _ = QFileDialog.getOpenFileName(
            None,
            "请选择Excel文件",
            self.settingdict["root_directory"],
            "Excel文件 (*.xlsx)"
        )
        # print(file_path)
        if file_path != '':                                                         # 选定了文件
            self.ui.fileChoose.setText(file_path)
            self.filePath = file_path

            self.__refresh_Sheet_name(file_path)                                    # 刷新Sheet界面
        else:                                                                       # 没有选定文件
            self.ui.fileChoose.setText("点击浏览根目录 . . .")


    def __close_check(self):                                                        # 最终校验函数
        self.isFileOK = False
        self.isSheetOK = False
        self.isColOK = False
        self.isNameOK = False
        if self.filePath != "":                                                     # 检查文件或者文件夹是否选错
            if self.filePath[-5:] == '.xlsx':
                # print(self.filePath[-5:])
                self.isFileOK = True
                                                           
        self.isSheetOK = True                                   # 检查Sheet是否选择， 只要文件选择正常不会出问题，已经默认选择第一个了

        self.ColName = self.ui.ColInput.text()                  # 校验姓名等的输入
        self.RowStart = self.ui.RowStartInput.text()
        self.RowEnd = self.ui.RowEndInput.text()
        if self.ColName != '' and self.RowStart != '' and self.RowEnd != '':
            try:
                correctCol = re.search(r'[A-Z]+', self.ColName).group()
                correctRowStart = re.search(r'[0-9]+', self.RowStart).group()
                correctRowEnd = re.search(r'[0-9]+', self.RowEnd).group()
                if correctCol == self.ColName and correctRowStart == self.RowStart and correctRowEnd == self.RowEnd:
                    self.RowStart = int(self.RowStart)
                    self.RowEnd = int(self.RowEnd)
                    if self.RowStart < self.RowEnd:
                        self.isColOK = True
            except Exception as e:
                self.isColOK = False

        self.ColScore = self.ui.ColScoreInput.text()            # 检验成绩的输入列
        try:
            if self.ColScore != '':
                # print(self.ColScore)
                correctScore = re.search(r'[A-Z]+', self.ColScore).group()
                if correctScore == self.ColScore:
                    self.isNameOK = True
        except Exception as e:
            self.isNameOK = False
        ## 最终检查
        if self.isFileOK and self.isSheetOK and self.isColOK and self.isNameOK:
            # print(self.returnValue())
            self.accept()
        else:
            self.__clock_check_prompt()

    def __clock_check_prompt(self):         # 用来提示有哪些选项没有选择
        prompt = "请输入或检查: "
        if self.isFileOK == False:
            prompt += "文件 "
        if self.isSheetOK == False:
            prompt += "Sheet "
        if self.isColOK == False:
            prompt += "学生姓名配置 "
        if self.isNameOK == False:
            prompt += "成绩输入配置 "
        self.ui.label.setText(prompt)



    ### ———————————————————————————— 初始化界面 ————————————————————————————
    def __init_all_uis(self):
        if self.settingdict["init_mode"] == 2:                                      # 默认选定文件夹
            self.ui.init_mode.setText("已经选定文件夹：")
            self.ui.file_path.setText(self.settingdict["root_directory"])
            self.ui.fileChoose.setText("点击浏览根目录 . . .")
        else:                                                                       # 默认选定文件
            self.ui.init_mode.setText("已经选定文件：")
            self.ui.file_path.setText(self.settingdict["file_name"])
            self.ui.fileChoose.setText("文件已被选定，不可更改")
            self.ui.fileChoose.setEnabled(False)
            if self.settingdict["init_mode"] == 0:                                  # 默认选定Sheet
                self.ui.SheetChoose.addItem(self.settingdict["sheet_name"])
                self.ui.SheetMode.setText("当前Sheet已经选定，不可更改")
                self.ui.SheetChoose.setEnabled(False)
            else:
                self.__refresh_Sheet_name(self.filePath)

        self.ui.ColInput.setPlaceholderText("请输入文本 (如 大写A,B,C . . .)")
        self.ui.RowStartInput.setPlaceholderText("请输入整数 (如 1,3,5 . . .)")
        self.ui.RowEndInput.setPlaceholderText("请输入整数 (如 57,72 . . .)")
        if self.settingdict["NAME_Init_mode"] == 0:                                 # 默认选定姓名的行列
            self.ui.ColInput.setText(self.settingdict["NAME_COL"])
            self.ui.RowStartInput.setText(str(self.settingdict["NAME_ROW_START"]))
            self.ui.RowEndInput.setText(str(self.settingdict["NAME_ROW_END"]))
        
        self.ui.ColScoreInput.setPlaceholderText("请输入文本 (如 大写A,B,C . . .)")



    ## ---------------------------- Sheet刷新和选择 ----------------------------
    def __refresh_Sheet_name(self, file_path:str):
        all_Sheet = mp.find_xlsx_Sheets(file_path)
        # print(all_Sheet)
        self.ui.SheetChoose.clear()
        # length = len(all_Sheet)
        self.ui.SheetChoose.addItems(all_Sheet)

    ### ———————————————————————————— 对外接口 ————————————————————————————
    def returnValue(self)->dict:
        self.SheetName = self.ui.SheetChoose.currentText()
        return {
            "FileName" : self.filePath,
            "SheetName" : self.SheetName,
            "ColName" : self.ColName,
            "RowStart" : self.RowStart,
            "RowEnd" : self.RowEnd,
            "ColScore" : self.ColScore
        }



if __name__ == "__main__":
    app = QApplication()
    window = FileNameSettingsWindow({'root_directory': 'D:/Programme/python/ReadAndRecord/Excel', 
                                     'file_name': '00.xlsx', 
                                     'whole_file_path': 'D:/Programme/python/ReadAndRecord/Excel/00.xlsx', 
                                     'sheet_name': 'Sheet2', 
                                     'NAME_Init_mode': 0, 
                                     'NAME_COL': 'C', 
                                     'NAME_ROW_START': 2, 
                                     'NAME_ROW_END': 99, 
                                     'init_mode': 2, 
                                     'usingAudioRecognition': False})
    
    # window = FileNameSettingsWindow({'root_directory': 'D:/Programme/python/ReadAndRecord', 
    #                                  'file_name': '', 
    #                                  'whole_file_path': 'D:/Programme/python/ReadAndRecord', 
    #                                  'sheet_name': '', 
    #                                  'NAME_Init_mode': 0, 
    #                                  'NAME_COL': 'C', 
    #                                  'NAME_ROW_START': 3, 
    #                                  'NAME_ROW_END': 99, 
    #                                  'init_mode': 2, 
    #                                  'usingAudioRecognition': False})

    window.show()
    window.statusTip()

    app.exec()