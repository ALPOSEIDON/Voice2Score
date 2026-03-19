import re

def cmd_analyse_name_score(cmd:str):
    # 正则表达式还需要修改
    m = re.search(r"(?P<name>[\u4e00-\u9fff]+)\s*(?P<score>\d+\.?5?)", cmd)
    return m

def score_analyze(text:str)->str:
    m = re.search(r"[\u4e00-\u9fff]*\s*(?P<score>\d+\.?5?)", text)
    if m != None:
        return m["score"]
    return None



if __name__ == "__main__":
    print(score_analyze("阿斯蒂芬   11.5") + "11")