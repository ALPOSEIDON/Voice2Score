from PySide6.QtWidgets import QMessageBox, QHeaderView, QFileDialog, QDialog, QApplication
from PySide6.QtUiTools import QUiLoader
from PySide6.QtCore import Signal,QObject, Slot, QThread
import re
import MyPackages as mp


class DefaultSettingsWindow(QDialog):

    returnValueTuple = (
            "init_mode" ,
            "FileName" ,
            "SheetName",
            "Col_Choice",
            "ColName",
            "RowStart",
            "RowEnd",
            "isAudio_recognizer_init",
        )

    def __init__(self):
        super().__init__()
        try:
            # 返回的对象是一个 QMainWindow类的实例 也就是我创建的主窗口
            self.ui = QUiLoader().load(r'Resources\backup\DefaultSettings.ui')
        except Exception as e:
            QMessageBox.critical(None, "初始化异常", 
                                 f"读取失败：{type(e).__name__}\n{e}\n请检查 Resources\\backup\DefaultSettings.ui 文件是否出现问题")
        
        # 把 UI 的布局加载进去
        lay = self.ui.layout()
        self.setLayout(lay)
        self.resize(480, 462)
        self.move(300, 200)
        self.setWindowTitle("默认设置")

        # 先把一些参数初始化
        self.c_init_mode = 0
        self.c_ColChoose = 0                        # 默认是否指定学生姓名所在的列

        self.isFileOK = False
        self.isSheetOK = False
        self.isColOK = False

        self.filePath = ""                          # 选择的文件路径
        self.SheetName = ""                         # 选择的Sheet名
        self.ColName = ""                           # 选择的Col名
        self.RowStart = 0
        self.RowEnd = 0
        self.isAudio_recognizer_init = False        # 默认语音识别是否开启

        self.ui.buttonBox.accepted.connect(self.__close_check)  # 可以连接到我自己的函数上，方便self.accept()自动调用
        self.ui.buttonBox.rejected.connect(self.reject)
        self.ui.init_modeChoose.currentIndexChanged.connect(self.__init_mode_combo) # activated
        self.ui.ColChoose.activated.connect(self.__ColChoose_combo)
        self.ui.audio_recognizer.clicked.connect(self.__audio_recognizer_check)
        self.ui.fileChoose.clicked.connect(self.__fileChoosing_dialog)

        # 对界面进行初始化
        self.__init_mode0_combo()


    ### ———————————————————————————— 信号处理函数 ————————————————————————————
    ## ---------------------------- 默认文件夹选择模式 ----------------------------
    def __init_mode_combo(self):
        # print("ok")
        ## 刷新界面
        self.ui.fileChoose.setText("点击浏览 . . .")
        self.ui.SheetChoose.clear()
        self.ui.ColInput.setText("")
        self.ui.RowStartInput.setText("")
        self.ui.RowEndInput.setText("")
        # 0:指定Excel文件和Sheet 1:仅指定Excel文件 2:仅指定文件夹(根目录)
        currentChoice = self.ui.init_modeChoose.currentIndex()
        self.c_init_mode = currentChoice
        if currentChoice == 0:
            self.__init_mode0_combo()
        elif currentChoice == 1:
            self.__init_mode1_combo()
        else:
            self.__init_mode2_combo()

    def __init_mode0_combo(self):
        self.ui.SheetChoose.setEnabled(True)
        self.ui.ColChoose.setCurrentIndex(0)
        self.c_ColChoose = 0
        self.StudentNameLineEdit(True)

    def __init_mode1_combo(self):
        self.ui.SheetChoose.setEnabled(False)
        self.ui.ColChoose.setCurrentIndex(0)
        self.c_ColChoose = 0
        self.StudentNameLineEdit(True)

    def __init_mode2_combo(self):
        self.ui.SheetChoose.setEnabled(False)
        self.ui.ColChoose.setCurrentIndex(1)
        self.c_ColChoose = 1
        self.StudentNameLineEdit(False)


    def StudentNameLineEdit(self, status:bool):
        if status == True:
            self.ui.ColInput.setEnabled(True)
            self.ui.RowStartInput.setEnabled(True)
            self.ui.RowEndInput.setEnabled(True)
            self.ui.ColInput.setPlaceholderText("请输入文本 (如 大写A,B,C . . .)")
            self.ui.RowStartInput.setPlaceholderText("请输入整数 (如 1,3,5 . . .)")
            self.ui.RowEndInput.setPlaceholderText("请输入整数 (如 57,72 . . .)")
        else:
            self.ui.ColInput.setEnabled(False)
            self.ui.RowStartInput.setEnabled(False)
            self.ui.RowEndInput.setEnabled(False)
            self.ui.ColInput.setPlaceholderText("当前选项已被禁用，您无需输入任何数据")
            self.ui.RowStartInput.setPlaceholderText("当前选项已被禁用，您无需输入任何数据")
            self.ui.RowEndInput.setPlaceholderText("当前选项已被禁用，您无需输入任何数据")


    ## ---------------------------- 选择文件 ----------------------------
    def __fileChoosing_dialog(self):
        if self.c_init_mode == 0 or self.c_init_mode == 1:
            file_path, _ = QFileDialog.getOpenFileName(
                None,
                "请选择Excel文件",
                "./Excel",
                "Excel文件 (*.xlsx)"
            )
            # print(file_path)
            if file_path != '':
                self.ui.fileChoose.setText(file_path)
                self.filePath = file_path

                if self.c_init_mode == 0:
                    self.__refresh_Sheet_name(file_path)                # 刷新Sheet界面
            else:
                self.ui.fileChoose.setText("点击浏览 . . .")
        else:
            folder = QFileDialog.getExistingDirectory(
                None,
                "请选择工作区根目录",
                "."
            )
            # print(folder == '')
            # print(folder)
            if folder != '':
                self.ui.fileChoose.setText(folder)
                self.filePath = folder
            else:
                self.ui.fileChoose.setText("点击浏览 . . .")


    ## 是否选定学生 所在的的列
    def __ColChoose_combo(self):
        currentChoice = self.ui.ColChoose.currentIndex()
        self.c_ColChoose = currentChoice
        if currentChoice == 0:
            self.StudentNameLineEdit(True)
        else:
            self.StudentNameLineEdit(False)
        

    ## 是否默认打开语音识别
    ## 语音提示
    def __audio_recognizer_check(self):
        if self.ui.audio_recognizer.isChecked():
            self.isAudio_recognizer_init = True
        else:
            self.isAudio_recognizer_init = False
        # print(self.isAudio_recognizer_init)

    ## ---------------------------- Sheet刷新和选择 ----------------------------
    def __refresh_Sheet_name(self, file_path:str):
        all_Sheet = mp.find_xlsx_Sheets(file_path)
        # print(all_Sheet)
        self.ui.SheetChoose.clear()
        # length = len(all_Sheet)
        self.ui.SheetChoose.addItems(all_Sheet)

    ### ———————————————————————————— 关闭信号等处理 ————————————————————————————
    def __close_check(self):
        # 检查选项初始化
        self.isFileOK = False
        self.isSheetOK = False
        self.isColOK = False
        # 开始检查各个选项
        if self.filePath != "":             # 检查文件或者文件夹是否选错
            if self.c_init_mode == 0 or self.c_init_mode == 1: # 选择的是文件
                if self.filePath[-5:] == '.xlsx':
                    # print(self.filePath[-5:])
                    self.isFileOK = True
            else:
                if self.filePath[-5:] != '.xlsx':
                    self.isFileOK = True
        if self.c_init_mode == 0:           # 检查Sheet是否选择， 只要文件选择正常不会出问题，已经默认选择第一个了
            self.isSheetOK = True
        else:
            self.isSheetOK = True           # 如果不用Sheet直接通过，留下空间以供未来检验
        if self.c_ColChoose == 0:           # 检查Col是否正确输入
            self.ColName = self.ui.ColInput.text()
            self.RowStart = self.ui.RowStartInput.text()
            self.RowEnd = self.ui.RowEndInput.text()
            if self.ColName != '':
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
        else:                               # 如果禁用Col，直接通过校验
            self.isColOK = True
        # 最终检查
        # print(self.isFileOK, self.isSheetOK, self.isColOK)
        if self.isFileOK and self.isSheetOK and self.isColOK:
            self.accept()
        else:
            self.__clock_check_prompt()

    def __clock_check_prompt(self):         # 用来提示有哪些选项没有选择
        prompt = ""
        if self.isFileOK == False:
            prompt += "请选择文件 "
        if self.isSheetOK == False:
            prompt += "请选择Sheet "
        if self.isColOK == False:
            prompt += "请输入或检查Col "
        self.ui.prompt_label.setText(prompt)

    

    ### ———————————————————————————— 对外接口 ————————————————————————————
    def returnValue(self)->dict:
        if self.c_init_mode == 0:
            self.SheetName = self.ui.SheetChoose.currentText()
        else:
            self.SheetName = ''
        # print(self.SheetName == '')
        return {
            "init_mode" : self.c_init_mode,
            "FileName" : self.filePath,
            "SheetName" : self.SheetName,
            "Col_Choice" : self.c_ColChoose,
            "ColName" : self.ColName,
            "RowStart" : self.RowStart,
            "RowEnd" : self.RowEnd,
            "isAudio_recognizer_init" : self.isAudio_recognizer_init
        }
            

if __name__ == "__main__":
    app = QApplication()
    window = DefaultSettingsWindow()

    window.show()
    window.statusTip()

    app.exec()