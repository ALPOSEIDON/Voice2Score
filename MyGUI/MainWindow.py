from PySide6.QtWidgets import QMessageBox, QHeaderView, QMainWindow, QTableWidgetItem
from PySide6.QtGui import QColor
from PySide6.QtUiTools import QUiLoader
from PySide6.QtCore import Signal,QObject, Slot, QThread, Qt
from .ChildWindow import DefaultSettingsWindow, FileNameSettingsWindow
import re
import MyPackages as mp

# 默认配置 名称的注册表
_DEFAULT_SETTINGS_MENU = (
    "root_directory", 
    "file_name",
    "whole_file_path",
    "sheet_name",
    "NAME_Init_mode",  # 0：确定初始化 1：确定不初始化
    "NAME_COL",
    "NAME_ROW_START",
    "NAME_ROW_END",
    "init_mode",    # 用于指定初始化方式： 0：直接打开路径中的文件和对应的Sheet， 1：指定excel，需要指定文件 2：指定根目录，需要指定文件
    "usingAudioRecognition",
    "ColScore",
    )
# settings 文件位置
_SETTINGS_STORAGER = r'Resources\backup\settings.json'
_MODEL = r"Resources\vosk-model-cn-0.22"

_DEFAULT_COLUMN_COUNT = 5

 
class MWindow(QMainWindow):
    """
    MWindow 的 Docstring

    内部的主要参数有 settings audio textOutput cmdInput
    """
    def __init__(self):
        super().__init__()
        # 从文件中加载UI定义
 
        # 从 UI 定义中动态 创建一个相应的窗口对象
        # 注意：里面的控件对象也成为窗口对象的属性了
        # 比如 self.ui.button , self.ui.textEdit

        ### UI初始化
        try:
            # 返回的对象是一个 QMainWindow类的实例 也就是我创建的主窗口
            self.ui = QUiLoader().load(r'Resources\backup\MyMain.ui')
        except Exception as e:
            QMessageBox.critical(None, "初始化异常", 
                                 f"读取失败：{type(e).__name__}\n{e}\n请检查 Resources\\backup\MyMain.ui 文件是否出现问题")
            exit(-1)
        self.ui.debugOutput.appendPlainText(' UI 文件加载成功')
        self.ui.nameWidget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        # self.ui.nameWidget.horizontalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.setCentralWidget(self.ui)
        self.resize(880, 700)
        self.move(300, 200)
        self.setWindowTitle("语音录入小程序")
        self.ui.cancelInput.setEnabled(False)
        self.ui.nameWidget.setEnabled(False)



        ### 事件绑定
        ## vosk 线程初始化
        self.threadVosk = QThread(self)
        self.worker = voskWorker()
        self.worker.moveToThread(self.threadVosk)
        self.worker.initOK.connect(self.__vosk_init_ok)
        self.worker.initFail.connect(self.__vosk_init_failure)
        self.worker.RecognizedName.connect(self.__score_handle_display)
        self.threadVosk.start()
        # self.worker.initModel.emit(_MODEL)

        # self.ui.leftNumbers.display(55)

        ## UI 界面信号绑定
        self.ui.default_settings.clicked.connect(self.__settings_clicked)
        # self.ui.audio_prompt.clicked.connect(self.__audio_check_status)               # 打开语音提示，这个功能暂时关闭
        self.ui.speech_recognition.clicked.connect(self.__speech_recognition_check_status)
        # self.ui.cmdInput.returnPressed.connect(self._cmdInput_check_cmd)              # 在line edit里边编辑，被我移除了
        self.ui.pressedToStart.clicked.connect(self.__entering_score_check)
        self.ui.cancelInput.clicked.connect(self.__cancelInput_check)
        self.ui.fileChooseFinal.clicked.connect(self.__fileChooseFinal_check)
        self.ui.nameWidget.itemChanged.connect(self.__Item_name_changing)



        # 加载默认设置
        try:
            self.settingdict = mp.file_loader(_SETTINGS_STORAGER)
        except Exception as e:
            QMessageBox.critical(None, "默认设置加载异常", 
                                 f"读取失败：{type(e).__name__}\n{e}\n请检查 Resources\\backup\\settings.json 文件是否出现问题")
            exit(-1)

        # 检查默认设置
        try:
            if len(_DEFAULT_SETTINGS_MENU) != len(self.settingdict):
                raise RuntimeError("")
            for key, value in self.settingdict.items():
                _DEFAULT_SETTINGS_MENU.index(key)
        except Exception:
            QMessageBox.critical(None, "默认设置加载异常", 
                                 f"读取失败\n请检查 Resources\\backup\\settings.json 文件是否出现问题")
            exit(-1)

        ### 类成员变量初始化
        self.__isAudioCreated = 0
        self.__usingVoskModel = False
        self.__InitOK = 0
        self.cmdInput_check_mode = 1    # 用来检查按下回车后应该干啥 -1：未初始化 0：输入sheet等等进行初始化 1：开始录入成绩 
        self.__needToRefresh = True
        self.__isSettingsAllOK = False
        # 初始化模式 0：直接打开路径中的文件和对应的Sheet， 1：指定excel，需要指定Sheet 2：指定根目录，需要指定文件 3:什么都没指定
        self.init_mode = self.settingdict["init_mode"]      
        self.init_mode_now = self.init_mode
        # self.initializing(self.init_mode)

        # 加载姓名
        self.namelists = ''
        self.info_stack = ''                          # 初始化，提前占位
        self.__default_setting_load()
        self.MyExcel = ''                         # 初始化，提前占位
        ### ---------------- 测试用 ------------------------
        self.__InitOK = 1
        # self.__init_tableWidget()






    ### ———————————————————————————— 操作函数 ————————————————————————————
    ## 删除信息并且刷新
    def refresh_excel_display(self, name:str, score:str):
        name_index = self.info_stack.name_index(name)
        try:
            self.MyExcel.write_score(name_index, float(score))
        except Exception as e:
            self.ui.debugOutput.appendPlainText("请关闭录入成绩的Excel !!")
            self._name_score_display("关闭Excel", 0, False)
            QMessageBox.critical(None, "Excel加载异常", 
                    f"读取失败：{type(e).__name__}\n{e}\n请关闭正在录入成绩的 Excel")
            return False
        self.info_stack.delete_name(name)
        self._name_score_display(name, score, True)
        self.ui.debugOutput.appendPlainText(name + " 同学的成绩：" + score  + " 已经登入")
        return True




    ### ———————————————————————————— 控制信号 ————————————————————————————

    ### ———————————————————————————— 最终加载姓名和所有设置 ————————————————————————————
    def __all_namelists_refresh(self):
        if self.__needToRefresh == True:
            try:
                self.namelists = mp.file_loader(r"Resources\backup\namelists.json")
                self.info_stack = mp.name_heap(self.namelists, _DEFAULT_COLUMN_COUNT)
                self.__init_tableWidget()
                self.MyExcel = mp.ExcelIO(self.settingdict["whole_file_path"], 
                                        self.settingdict["sheet_name"], 
                                        self.settingdict["ColScore"], 
                                        self.settingdict["NAME_COL"], 
                                        self.settingdict["NAME_ROW_START"])
            except Exception as e:
                QMessageBox.critical(None, "默认设置加载异常", 
                                    f"读取失败：{type(e).__name__}\n{e}\n您的 Resources文件夹 似乎出现了错误")
                exit(-1)
            

            self.__needToRefresh = False



    ## ———————————————————————————— 开始录入成绩按键 ————————————————————————————
    def __entering_score_check(self):
        if self.ui.pressedToStart.isChecked():
            self.__start_entering_score()
        else:
            self.__end_entering_score()

    def __start_entering_score(self):
        self.__InitOK_check()
        if self.__InitOK == 1:
            # 将设置等选项禁用
            self.ui.speech_recognition.setEnabled(False)
            self.ui.default_settings.setEnabled(False)
            self.ui.fileChooseFinal.setEnabled(False)
            self.ui.cancelInput.setEnabled(True)
            self.ui.debugOutput.appendPlainText('配置成功，可以开始录入成绩!')
            self.ui.nameWidget.setEnabled(True)
            self.ui.leftNumbers.display(str(len(self.info_stack.current_namelist)))


            if self.__usingVoskModel:
                self.worker.vosk.runningVosk = True
                self.worker.RecognizingName.emit(self.info_stack.current_namelist, self.info_stack.current_PY_INDEX)
                self.worker.stillUsing = True
        else:
            self.ui.debugOutput.appendPlainText('参数未初始化完全，请指定文件！')
            self.ui.pressedToStart.setChecked(False)

    def __end_entering_score(self):
        if self.__usingVoskModel:
            self.worker.vosk.runningVosk = False
            self.worker.stillUsing = False
            # print(self.worker.vosk.runningVosk)
        self.ui.speech_recognition.setEnabled(True)
        self.ui.default_settings.setEnabled(True)
        self.ui.fileChooseFinal.setEnabled(True)

        # self.ui.nameWidget.setEnabled(False)
        # self.ui.cancelInput.setEnabled(False)

        self.ui.debugOutput.appendPlainText('已经暂停录入成绩')

    def __InitOK_check(self):
        self.__InitOK = 0
        if self.__isSettingsAllOK == False:
            self.__InitOK += 1
        else:
            if self.__needToRefresh == True:
                self.__all_namelists_refresh()
                if self.__needToRefresh == True:
                    self.__InitOK += 1

        if self.__InitOK != 0:
            self.__InitOK = 0
        else:
            self.__InitOK = 1




    ## ———————————————————————————— 默认设置配置和加载 ————————————————————————————
    # ---------------------------- 默认设置配置 ----------------------------
    def __settings_clicked(self):
        dlg = DefaultSettingsWindow()
        self.ui.debugOutput.appendPlainText('已经打开默认设置界面')
        if dlg.exec():
            # print(dlg.returnValue())
            # 保存默认配置
            newDefaultSettings = dlg.returnValue()
            self.settingdict["init_mode"] = newDefaultSettings["init_mode"]                     # 通用设置加载
            self.settingdict["whole_file_path"] = newDefaultSettings["FileName"]
            self.settingdict["sheet_name"] = newDefaultSettings["SheetName"]
            self.settingdict["NAME_Init_mode"] = newDefaultSettings["Col_Choice"]
            self.settingdict["NAME_COL"] = newDefaultSettings["ColName"]
            self.settingdict["NAME_ROW_START"] = newDefaultSettings["RowStart"]
            self.settingdict["NAME_ROW_END"] = newDefaultSettings["RowEnd"]
            self.settingdict["usingAudioRecognition"] = newDefaultSettings["isAudio_recognizer_init"]
            if self.settingdict["init_mode"] == 2:                                              # 个性化设置 只存储根目录
                self.settingdict["root_directory"] = self.settingdict["whole_file_path"]
                self.settingdict["file_name"] = ''
            else:                                                                               # 文件名也要存储
                searchANS = re.search(r"(?P<root>.+)/(?P<file>.+?\.xlsx)", self.settingdict["whole_file_path"]) # (?P<root>.+)\\(?P<file>.+?\.xlsx)  /.+?\.xlsx
                # print(searchANS["root"])
                # print(searchANS["file"])
                self.settingdict["root_directory"] = searchANS["root"]
                self.settingdict["file_name"] = searchANS["file"]

            self.settingdict["ColScore"] = ""
            mp.file_store(r'Resources\backup\settings.json', self.settingdict)
            self.ui.debugOutput.appendPlainText('您已成功配置默认设置并保存')

            self.__needToRefresh = True
            self.__isSettingsAllOK = False
            # print(self.settingdict)
        else:
            self.ui.debugOutput.appendPlainText('默认设置配置失败')

    # ---------------------------- 最终设置配置 ----------------------------
    def __fileChooseFinal_check(self):
        dlg = FileNameSettingsWindow(self.settingdict)
        self.ui.debugOutput.appendPlainText('已经打开文件设置界面')
        if dlg.exec():
            finalSettings = dlg.returnValue()
            # print(finalSettings)                     
            self.settingdict["whole_file_path"] = finalSettings["FileName"]                 # 通用设置加载
            self.settingdict["sheet_name"] = finalSettings["SheetName"]
            self.settingdict["NAME_COL"] = finalSettings["ColName"]
            self.settingdict["NAME_ROW_START"] = finalSettings["RowStart"]
            self.settingdict["NAME_ROW_END"] = finalSettings["RowEnd"]
            self.settingdict["ColScore"] = finalSettings["ColScore"]

            # 刷新一下名单，防止数据被篡改
            try:
                mp.refresh_namelists(self.settingdict["NAME_COL"],
                                    self.settingdict["NAME_ROW_START"],
                                    self.settingdict["NAME_ROW_END"],
                                    self.settingdict["whole_file_path"],
                                    self.settingdict["sheet_name"],
                                    r"Resources\backup\namelists.json")
            except Exception as e:
                QMessageBox.critical(None, "初始化异常", 
                                 f"读取失败：{type(e).__name__}\n{e}\n或许是您输入的姓名列数超出限制，或选错了列")
                exit(-1)
            # 刷新完名单需要重新加载所有设置
            self.__needToRefresh = True
            self.__isSettingsAllOK = True

            self.ui.debugOutput.appendPlainText('您已成功配置最终设置')
        else:
            self.ui.debugOutput.appendPlainText('文件设置配置失败')
    

    ### 加载默认设置
    def __default_setting_load(self):
        if self.settingdict["usingAudioRecognition"] == True:
            self.ui.speech_recognition.setChecked(True)
            self.__speech_recognition_check_status()



    def closeEvent(self, event):
        self.worker.stop()
        self.threadVosk.quit()
        # self.threadVosk.wait()
        self.threadVosk.terminate()
        event.accept()


    

    ### ———————————————————————————— 姓名表格的函数 ————————————————————————————
    ## ---------------------------- 初始化 ----------------------------
    def __init_tableWidget(self):       # 直接读取数据，如果有需要可以提前刷新 namelist 和 info_stack
        self.ui.nameWidget.blockSignals(True)                                                               # 阻塞信号
        self.ui.nameWidget.clear()                                                                          # 清空界面
        totalNum = len(self.info_stack.all_namelist)
        RowCount, ColCount = self.info_stack.query_location_in_table(self.info_stack.all_namelist[totalNum - 1])
        self.ui.nameWidget.setRowCount(RowCount + 1)
        self.ui.nameWidget.setColumnCount(_DEFAULT_COLUMN_COUNT)
        # print(totalNum)
        for i in range(totalNum):
            c_Row, c_Col = self.info_stack.query_location_in_table(self.info_stack.all_namelist[i])
            # print(self.info_stack.all_namelist[i])
            self.ui.nameWidget.setItem(c_Row, c_Col, QTableWidgetItem(self.info_stack.all_namelist[i]))
            # self.ui.nameWidget.item(c_Row, c_Col).setTextAlignment(Qt.AlignmentFlag.AlignHCenter)           # 数据居中放置
            self.ui.nameWidget.item(c_Row, c_Col).setTextAlignment(Qt.AlignmentFlag.AlignCenter)           # 数据居中放置
        self.ui.nameWidget.blockSignals(False)

    def __Item_name_changing(self, item):
        # print(1)
        self.ui.nameWidget.blockSignals(True)
        row = item.row()
        col = item.column()
        text = item.text()
        # print(text)
        nameIndex = row * _DEFAULT_COLUMN_COUNT + col
        score = mp.score_analyze(text)
        if nameIndex + 1 > len(self.info_stack.all_namelist):
            self.ui.debugOutput.appendPlainText('输入位置超出界限')
        elif score == None:
            self.ui.debugOutput.appendPlainText('您的输入检测不到成绩，请重试')
        else:
            name = self.info_stack.all_namelist[nameIndex]
            if self.refresh_excel_display(name, score) == True:
                # self.info_stack.delete_name(name)
                item.setText(name + ":" + score)
                
                ## 刷新状态栏
                item.setForeground(QColor(255, 255, 0))
                ## 锁定状态栏
                item.setFlags(Qt.ItemFlag.ItemIsEnabled)
        self.ui.nameWidget.blockSignals(False)
            





    ## 语音提示
    def __audio_check_status(self):
        if self.ui.audio_prompt.isChecked():
            self.__audio_init()
        else:
            self.__audio_deinit()

    def __audio_init(self):
        if self.__isAudioCreated == 0:
            self.ui.debugOutput.appendPlainText('正在初始化引擎')

            self.ui.debugOutput.appendPlainText('引擎初始化成功')
            self.__isAudioCreated = 1
        else:
            self.ui.debugOutput.appendPlainText('已启用之前的引擎')

    def __audio_deinit(self):
        self.ui.debugOutput.appendPlainText('引擎已关闭')


    ## 语音识别
    def __speech_recognition_check_status(self):
        if self.ui.speech_recognition.isChecked():
            self.__speech_recognition_init()
        else:
            self.__speech_recognition_deinit()

    def __speech_recognition_init(self):
        if self.worker.isInitializing:
            self.ui.debugOutput.appendPlainText('vosk 引擎正在初始化，请稍等!')
        elif self.worker.isInitialized:
            self.ui.debugOutput.appendPlainText('vosk 引擎之前已被初始化过，已复用!')
            self.__usingVoskModel = True
            self.worker.stillUsing = True
        else:
            self.ui.debugOutput.appendPlainText('vosk 引擎正在初始化，请稍等!')
            self.ui.speech_recognition.setEnabled(False)
            self.ui.pressedToStart.setEnabled(False)
            self.ui.studentName.setText("初始化中")
            self.worker.initModel.emit(_MODEL)

    def __speech_recognition_deinit(self):
        self.ui.debugOutput.appendPlainText('vosk 引擎已关闭!')
        self.__usingVoskModel = False
        self.worker.stillUsing = False

    def __vosk_init_ok(self):
        self.__usingVoskModel = True
        self.worker.stillUsing = True
        self.ui.debugOutput.appendPlainText('vosk 引擎初始化成功，您可以准备读成绩!')
        self.ui.speech_recognition.setEnabled(True)
        self.ui.pressedToStart.setEnabled(True)
        self.ui.studentName.setText("初始化完成")

    def __vosk_init_failure(self, str1, str2):
        QMessageBox.critical(None, "引擎初始化异常", 
                                 f"读取失败：{str1}\n{str2}\n请检查 Resources\\vosk-model-cn-0.22 文件是否出现问题")
        exit(-1)

    # ---------------------------- 返回信号处理 ----------------------------
    def __score_handle_display(self, name:str, score:str):
        if self.worker.stillUsing:
            if self.refresh_excel_display(name, score) == True:
                ## 修改Table状态
                self.ui.nameWidget.blockSignals(True)
                row, col = self.info_stack.query_location_in_table(name)
                item = self.ui.nameWidget.item(row, col)
                item.setText(name + ":" + score)
                
                    # 刷新状态栏
                item.setForeground(QColor(255, 255, 0))
                    # 锁定状态栏
                item.setFlags(Qt.ItemFlag.ItemIsEnabled)
                self.ui.nameWidget.blockSignals(False)

            self.worker.RecognizingName.emit(self.info_stack.current_namelist, self.info_stack.current_PY_INDEX)
    # 初始时信号发送 别忘了

    # 撤回输入按键 在开始输入按键没有按下的时候禁用 ！！
    def __cancelInput_check(self):
        length = len(self.info_stack.backup_namelist)
        if length <= 0:
            self.ui.debugOutput.appendPlainText('之前没有录入需要回溯!')
            self._name_score_display("没有历史", "0", False)
        else:
            name, score = self.info_stack.recall_name()
            try:
                self.MyExcel.delete_score(self.info_stack.name_index(name)) 
            except Exception as e:
                self.ui.debugOutput.appendPlainText("请关闭录入成绩的Excel !!")
                self._name_score_display("关闭Excel", 0, False)
                QMessageBox.critical(None, "Excel加载异常", 
                        f"读取失败：{type(e).__name__}\n{e}\n请关闭正在录入成绩的 Excel")
                self.info_stack.delete_name(name)
                return False
            self.ui.debugOutput.appendPlainText(name + " 同学的成绩 已经撤销")
            # 清除原本的Table禁用状态
            self.ui.nameWidget.blockSignals(True)
            row, col = self.info_stack.query_location_in_table(name)
            item = self.ui.nameWidget.item(row, col)
            item.setFlags(item.flags()|Qt.ItemFlag.ItemIsEditable)
            item.setText(name)
            item.setForeground(QColor(255, 255, 255))
            self.ui.nameWidget.blockSignals(False)
            
            self._name_score_display("撤回完成", "0", True)

            





    def _cmdInput_check_cmd(self):
        currentText = self.ui.cmdInput.text()
        if self.cmdInput_check_mode == 1:
            if currentText == '':
                self.ui.debugOutput.appendPlainText('语音输入开启')
            else:
                ans = mp.cmd_analyse_name_score(currentText)
                if ans != None:     # 检测到正确的输入
                    name = str(ans['name'])
                    score = str(ans['score'])
                    self.ui.debugOutput.appendPlainText(name + " 同学  成绩：" + score + " 已经录入")
                    self.ui.cmdInput.clear()
                    self.ui.studentName.setText(name)
                    self.ui.studentScore.display(float(score))
                else:
                    self.ui.debugOutput.appendPlainText("输入错误，请检查姓名格式等")
        elif self.cmdInput_check_mode == 0:
            self.initializing(self.init_mode_now)


    

    def initializing(self, init_mode):
        if init_mode == 3:
            self.ui.debugOutput.appendPlainText("请输入所需要录入成绩的列：")
            # try:
            #     self.excel = mp.ExcelIO(self.settingdict["whole_file_path"], self.settingdict["sheet_name"])
            self.ui.debugOutput.appendPlainText(f'已经加载{self.settingdict["file_name"]}：{self.settingdict["sheet_name"]}')
        elif init_mode == 2:
            pass
        elif init_mode == 1:
            pass
        elif init_mode == 0:
            pass

    ## 使用 lcd 和 显示器 刷新 
    def _name_score_display(self, name:str, score:str, needNum:bool):
        self.ui.studentName.setText(name)
        self.ui.studentScore.display(float(score))
        if needNum == True:
            self.ui.leftNumbers.display(str(len(self.info_stack.current_namelist)))
        

class voskWorker(QObject):
    # 控制信号 - 外部调用Worker
    initModel = Signal(str)
    RecognizingName = Signal(list, dict)

    # 状态信号 - Worker向外报告
    initOK = Signal()
    initFail = Signal(str, str)
    RecognizedName = Signal(str, str)

    def __init__(self):
        super().__init__()
        self._running = True
        self.isInitializing = False
        self.isInitialized = False
        self.stillUsing = False
        
        self.initModel.connect(self.load_model)
        self.RecognizingName.connect(self.recognizing_with_model)

        

    @Slot(str)
    def load_model(self, model_path):
        self.isInitializing = True
        try:
            self.vosk = mp.voskREC(model_path)
        except Exception as e:
            self.initFail.emit(f"{type(e).__name__}", f"{e}")
            exit(-1)
        self.isInitializing = False
        self.isInitialized = True
        self.initOK.emit()

    @Slot(list, dict)
    def recognizing_with_model(self, c_namelist, c_PY_INDEX):
        if self.stillUsing:
            q_name, q_score = self.vosk.video_capture(c_namelist, c_PY_INDEX)
            self.RecognizedName.emit(q_name, q_score)

    

    def stop(self):
        self._running = False