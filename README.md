# Voice2Score
This is a small program with which you can enter scores to an excel document easily through reading it loudly .It can only support Chinese and have some bugs to be fixed. I just develop this program for my mom who is a hard-working teacher, and maybe next winter holiday I will update a new version.


这个小项目是作者在寒假心血来潮做的，顺便借此机会学习一下Qt6和语音识别，同时也是因为我妈（一个小学老师）天天让我帮她录成绩，索性就做了这样一个项目给她用(doge)。

## 实现功能
- 能够通过语音模型识别你读出的学生成绩并且记录在EXCEL文件中（虽然有的时候识别起来真的很慢）
- 能够在一个表格中输入成绩并同样登记到EXCEL中（防止语音识别不清楚，同时让成绩登记可视化）
- 将全部调试信息打印到控制台，方便实时监管成绩登入情况

## 如何使用
- 需要你在Vosk官网下载 vosk-model-cn-0.22 这个模型（那个简化版的似乎也可以试试，不过记得修改模型文件名）
- 确保运行环境中有PySide6, vosk, pyaudio, pypinyin, audioop, collections等库
