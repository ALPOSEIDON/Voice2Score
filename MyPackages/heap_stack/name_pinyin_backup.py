import pypinyin

class name_heap:
    def __init__(self, your_namelist:list, colCount:int = 5):
        # 这两个是常量，  不许动！
        self.all_namelist = your_namelist.copy()
        self.all_PY_INDEX = {n: ''.join(pypinyin.lazy_pinyin(n)) for n in your_namelist}
        self.colCount = colCount
        # 当前的数据
        self.current_namelist = your_namelist.copy()
        self.current_PY_INDEX = self.all_PY_INDEX.copy()
        # 用来备份和恢复的数据
        self.backup_namelist = []
        self.backup_PY_INDEX = {}

    

    def delete_name(self, name:str):
        if len(self.current_namelist) <= 0:
            return -1
        i = 0
        for ch in self.current_namelist:
            if ch == name:
                # i 为当前的索引
                cname = self.current_namelist[i]
                cpinyin = self.all_PY_INDEX[cname]

                self.backup_namelist.append(cname)
                self.backup_PY_INDEX.update({cname:cpinyin})

                self.current_namelist.pop(i)
                self.current_PY_INDEX.pop(cname)
                return 1
            i += 1
        
        return -2
    
    def recall_name(self):
        if len(self.backup_namelist) <= 0:
            return None, None
        name = self.backup_namelist.pop()
        self.backup_PY_INDEX.pop(name)
        pinyin = self.all_PY_INDEX[name]

        self.current_namelist.append(name)
        self.current_PY_INDEX.update({name:pinyin})
        return name, pinyin        

    def query_location_in_table(self, name:str):
        i = 0
        for ch in self.all_namelist:
            if ch == name:
                currentCol = i % self.colCount
                currentRow = int(i / self.colCount)
                # 采用 0 开头的方式
                return currentRow, currentCol
            i += 1

        return -1, 1
    
    def name_index(self, name:str):
        i = 0
        for ch in self.all_namelist:
            if ch == name:
                return i
            i += 1
        return -1
    
    def current_backup(self):
        le = len(self.backup_namelist)
        if le <= 0:
            return None, None
        name = self.backup_namelist[le - 1]
        return name





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
    
    mytst = name_heap(namelist)
    print(mytst.all_namelist)
    print(mytst.all_PY_INDEX)
    mytst.delete_name('陈晓瑞')
    print(mytst.current_namelist)
    q1, q2 = mytst.query_location_in_table('董锘')
    print(q1, q2)
    print(mytst.all_PY_INDEX)
    mytst.recall_name()
    print(mytst.current_namelist)
    print(mytst.current_PY_INDEX)