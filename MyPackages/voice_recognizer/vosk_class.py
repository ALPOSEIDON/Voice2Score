import vosk, pyaudio, collections, pypinyin, audioop, re, json
from Levenshtein import distance

MODEL = r"Resources\vosk-model-cn-0.22"


# vosk-recognizer
class voskREC:
    sound_keys = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "百", "点"]
    sound_key_py = {k: ''.join(pypinyin.lazy_pinyin(k)) for k in sound_keys}

    def __init__(self, model_path:str):
        self.model = vosk.Model(model_path)
        self.rec = vosk.KaldiRecognizer(self.model, 16000)
        self.p = pyaudio.PyAudio()
        self.stream = self.p.open(format=pyaudio.paInt16, channels=1, rate=16000, input=True, frames_per_buffer=8000)
        # stream_num = self.p.open(format=pyaudio.paInt16, channels=1,rate=16000, input=True, frames_per_buffer=800)
        # 定义一个 ring 容器用来采集信息， 计算音量
        self.ring = collections.deque(maxlen=35)            # 20
        self.ring_current = collections.deque(maxlen=8)     # 8
        # print(len(self.ring))
        self.runningVosk = False
    

    
    def recognize_all_each_call(self, namelist):
        print("开始识别！")
        nameVoice = []
        isnameVoiceOK = 0
        scoreVoice = []
        #print("请说出要查询的名字……")
        i = 1
        while True:
            data = self.stream.read(400, exception_on_overflow=False)   # 400/16000 = 1/40
            currentRING = audioop.rms(data, 2)
            self.ring.append(currentRING)
            self.ring_current.append(currentRING)

            if sum(self.ring_current) / len(self.ring_current) < 100:     # and currentRING < 100    ||  sum(self.ring) / len(self.ring) < 100
                #环境安静  and currentRING < 250
                print("name and score", len(nameVoice), len(scoreVoice))
                print(isnameVoiceOK)
                if len(nameVoice) > 10:
                    isnameVoiceOK = 1
                if scoreVoice:
                    # 姓名识别
                    self.rec.Reset()
                    self.rec.AcceptWaveform(b''.join(nameVoice))
                    q_name = json.loads(self.rec.FinalResult()).get("text", "")
                    q_name = re.sub(r'\s+', '', q_name)       # 返回纯正的字符串
                    print(q_name)
                    self.rec.Reset()

                    # 分数识别
                    self.rec.AcceptWaveform(b''.join(scoreVoice))
                    q_score_raw = json.loads(self.rec.Result())["text"].replace(" ", "")
                    print(q_score_raw)
                    # 3. 逐字比对（允许 1 个拼音编辑距离误差）
                    THRESH = 1
                    q_score = []
                    for ch in q_score_raw:
                        if ch in self.sound_keys:  # 完全匹配
                            q_score.append(ch)
                            continue
                        # 拼音容错
                        ch_py = ''.join(pypinyin.lazy_pinyin(ch))
                        best, best_d = None, 999
                        for k, k_py in self.sound_key_py.items():
                            d = distance(ch_py, k_py)
                            if d < best_d:
                                best, best_d = k, d
                        if best_d <= THRESH:
                            q_score.append(best)
                    self.rec.Reset()
                    if q_name and q_score: 
                        for ch in nameVoice:
                            print(audioop.rms(ch, 2))

                        print("____________________________________")
                        for ch in scoreVoice:
                            print(audioop.rms(ch, 2))


                        break
                    nameVoice.clear()
                    scoreVoice.clear()
                    print("清零")
                    i = 1
                    isnameVoiceOK = 0
            else:
                print(f"录入{i}次")
                i += 1
                if isnameVoiceOK == 0 and sum(self.ring_current) / len(self.ring_current) > 20:  # 参数 注意！
                    nameVoice.append(data)
                elif isnameVoiceOK == 1:
                    scoreVoice.append(data)
                else:
                    nameVoice.append(data)
                    isnameVoiceOK = 1
        # 模糊匹配
        # if not namelist:
        #     raise RuntimeError("姓名名单为空，请检查 names.txt 是否正确加载！")
        qpinyin = pypinyin.lazy_pinyin(q_name)
        qpy = ''.join(qpinyin)
        q_name = min(namelist, key=lambda n: 0.7 * distance(qpy, PY_INDEX[n]) + 0.3 * distance(q_name, n))
        return q_name, q_score
    

    def video_capture(self, namelist:list, PY_INDEX:dict):
        print("开始识别！")
        # 清空缓冲
        self.rec.Reset()
        self.ring.clear()
        frames_in_buffer = self.stream.get_read_available()
        if frames_in_buffer:
            self.stream.read(frames_in_buffer, exception_on_overflow=False)

        voice = []
        i = 1
        isOK = 0
        q_name = ""
        q_score = ""
        #print("请说出要查询的名字……")
        while self.runningVosk:
            data = self.stream.read(800, exception_on_overflow=False)   # 800/16000 = 1/20
            currentRING = audioop.rms(data, 2)
            self.ring.append(currentRING)

            if sum(self.ring) / len(self.ring) < 300 and currentRING < 20:         #环境安静
                if voice:
                    for ch in voice:
                        print(audioop.rms(ch, 2))
                    print("____________________________________")
                    for f in range(len(voice)):
                        if audioop.rms(voice[f], 2) <= 3:
                            q_name, q_score = self.name_score_recognize(voice[0:f], voice[f:], namelist, PY_INDEX)
                            q_score = self.scorelist_to_str(q_score)
                            # 需要判别
                            isOK = 1
                            print("识别成功!")
                            break
                    if isOK:
                        break
                    else:
                        voice.clear()
            else:
                print(f"录入{i}次")
                i += 1
                voice.append(data)


        if self.runningVosk == False:
            self.rec.Reset()
        # 返回值都是 str, str
        return q_name, q_score
    

    def name_score_recognize(self, nameVoice, scoreVoice, namelist:list, PY_INDEX:dict):
        ## 姓名识别
        self.rec.Reset()
        self.rec.AcceptWaveform(b''.join(nameVoice))
        q_name_raw = json.loads(self.rec.FinalResult()).get("text", "")
        q_name_raw = re.sub(r'\s+', '', q_name_raw)       # 返回纯正的字符串
        print(q_name_raw)
        qpinyin = pypinyin.lazy_pinyin(q_name_raw)
        qpy = ''.join(qpinyin)
        q_name = min(namelist, key=lambda n: 0.7 * distance(qpy, PY_INDEX[n]) + 0.3 * distance(q_name_raw, n))

        ## 分数识别
        self.rec.Reset()
        self.rec.AcceptWaveform(b''.join(scoreVoice))
        q_score_raw = json.loads(self.rec.Result())["text"].replace(" ", "")
        print(q_score_raw)
        # 逐字比对（允许 1 个拼音编辑距离误差）
        THRESH = 1
        q_score = []
        for ch in q_score_raw:
            if ch in self.sound_keys:  # 完全匹配
                q_score.append(ch)
                continue
            # 拼音容错
            ch_py = ''.join(pypinyin.lazy_pinyin(ch))
            best, best_d = None, 999
            for k, k_py in self.sound_key_py.items():
                d = distance(ch_py, k_py)
                if d < best_d:
                    best, best_d = k, d
            if best_d <= THRESH:
                q_score.append(best)

        ## 识别完毕
        self.rec.Reset()
        return q_name, q_score
    


    def scorelist_to_str(self, score:list)->str:
        num = 0
        key = 0
        i = 1
        for w in score:
            if w == ' ':
                continue
            if key == 1:
                if w in self.sound_keys[:10]:
                    num += self.sound_keys.index(w) * 0.1 ** i
                    i += 1
                else:
                    # print("""something went wrong 
                    #       错误代码01
                    #       小数点之后的数似乎不应该这么读""")
                    pass
            elif w in self.sound_keys[:10]:
                num += self.sound_keys.index(w)
            elif w == self.sound_keys[10]:
                num *= 10
            elif w == self.sound_keys[11]:
                num *= 100
            elif w == self.sound_keys[12]:
                key = 1
            else:
                pass
        return str(num)

        
    




    



if __name__ == "__main__":
    namelist = ['柏宇涵', '陈晓瑞', '陈星潮', '董锘', '董馨妍', '樊琪乐', 
                '范羽晨', '高俊熙', '高雨欣', '高祯怿', '胡晨辉', '李昌骏', 
                '李梦琪', '林晨曦', '林盈萱', '林子凡', '刘晗笑', '刘诗羽', 
                '刘耀泽', '穆维琦', '潘岳', '钱纯熙', '钱思媛', '邵彦鸣', 
                '王康宇', '王亮', '王梦文', '王然', '王思远', '王祥昊', 
                '王欣怡', '王雨涵', '王誉霏', '王子怡', '魏欣悦', '魏玉豪', 
                '魏子欣', '相恒钰', '谢修冉', '徐楷芮', '杨淑雅', '尹星辰', 
                '张嘉豪', '张岚琋', '张译兮', '张正', '赵睿', '赵振廷', 
                '仲奕诺', '周子然', '周睿', '周思延', '周子睿', '朱可儿', 
                '陈子艺']
    
    audio = voskREC(MODEL)
    PY_INDEX = {n: ''.join(pypinyin.lazy_pinyin(n)) for n in namelist}
    while True:
        judge = input("请输入指令")
        if judge == " ":
            break
        # name, score = audio.recognize_all_each_call(namelist)
        q_name, q_score = audio.video_capture(namelist, PY_INDEX)
        print(q_name, q_score)
