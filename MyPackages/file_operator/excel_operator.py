import pandas as pd
from openpyxl import load_workbook
from .json_operator import file_store

def find_xlsx_files(string:str) -> list:
    """
    findfiles 的 Docstring
    
    :param string: 说明
    :type string: str
    :return: 返回所有 .xlsx 文件的名字列表
    :rtype: list

    用来寻找固定目录中的 .xlsx 等文件
    """
    from pathlib import Path
    folder = Path(string)  # 改成你的目录
    xlsx_files = list(folder.glob('*.xlsx'))  # 注意大小写
    xlsx_file_names = []
    for i in xlsx_files:
        str = i.name
        if str[:2] == "~$":
            pass
        else:
            xlsx_file_names.append(str)
    #print('共找到 xlsx:', len(xlsx_files))
    #for i in xlsx_files:
    #    print(i.name)
    return xlsx_file_names
#print(find_xlsx_files("Excel"))

def find_xlsx_Sheets(string) -> list:
    """
    find_xlsx_Sheets 的 Docstring
    
    :param string: 说明
    :return: 说明
    :rtype: list

    返回指定 xlsx 文件中的所有 Sheets
    """
    with pd.ExcelFile(string) as xls:
        all_sheets = xls.sheet_names
    return all_sheets
#print(find_xlsx_Sheets(r"Excel\00.xlsx"))

def refresh_namelists(NAME_COL:str, NAME_COL_START:int, NAME_COL_END:int, 
                      file_path_xlsx:str, sheet_name_xlsx:str, 
                      file_path_dst:str):
    """
    refresh_namelists 的 Docstring
    
    :param NAME_COL: 姓名所在的列
    :type NAME_COL: str
    :param NAME_COL_START: 姓名所在的行起点
    :type NAME_COL_START: int
    :param NAME_COL_END: 姓名所在的行终点
    :type NAME_COL_END: int
    :param file_path_xlsx: excel 文件路径
    :type file_path_xlsx: str
    :param sheet_name_xlsx: sheet 名字指定
    :type sheet_name_xlsx: str
    :param file_path_dst: 目标存储文件路径
    :type file_path_dst: str

    将 excel 文件中的所有姓名加载到 json 文件中
    """
    START_POINT = NAME_COL + str(NAME_COL_START)
    start_row, start_col = addr_to_indices(START_POINT)
    END_POINT = NAME_COL + str(NAME_COL_END)
    end_row, end_col = addr_to_indices(END_POINT)
    nrows = 1 + end_row - start_row
    df_write = pd.read_excel(file_path_xlsx, sheet_name=sheet_name_xlsx,
                             skiprows=start_row - 1, usecols=NAME_COL,
                             nrows=nrows)
    list_names = []
    list_names.append(df_write.columns[0])
    for i in range(nrows - 1):
        list_names.append(df_write.iloc[i, 0])
    # print(list_names)
    file_store(file_path_dst, list_names)

def addr_to_indices(cell_addr):
    """'B3' → (3, 2)  (row, col)"""
    from openpyxl.utils import range_boundaries
    c, r = range_boundaries(cell_addr + ':' + cell_addr)[:2]
    return r, c



class ExcelIO:
    def __init__(self, file_path:str, sheet_name:str, score_col:str,
                 NAME_COL:str, NAME_ROW_START:int):
        """
        __init__ 的 Docstring
        
        :param file_path: excel 文件路径
        :type file_path: str
        :param sheet_name: sheet 名字指定
        :type sheet_name: str
        :param score_col: 说明
        :type score_col: str
        :param NAME_COL: 姓名所在的列
        :type NAME_COL: str
        :param NAME_ROW_START: 姓名所在的行起点
        :type NAME_ROW_START: int
        """
        self.file_path = file_path
        self.book = load_workbook(file_path)
        self.file = self.book[sheet_name]
        self.score_col = score_col

        START_POINT = NAME_COL + str(NAME_ROW_START)
        self.start_row, self.start_col = addr_to_indices(START_POINT)

    def write_score(self, student_index:int, score:int|float):
        """
        write_score 的 Docstring
        
        :param student_index: 学生在原始数组中的下标
        :type student_index: int
        :param score: 学生成绩
        :type score: int | float
        """
        self.file[self.score_col + str(student_index + self.start_row)] = score
        self.book.save(self.file_path)
        # print("录入成功",name, score)

    def delete_score(self, student_index:int):
        """
        write_score 的 Docstring
        
        :param student_index: 学生在原始数组中的下标
        :type student_index: int
        :param score: 学生成绩
        :type score: int | float
        """
        self.file[self.score_col + str(student_index + self.start_row)] = ""
        self.book.save(self.file_path)
        # print("录入成功",name, score)

    def findSheets(self, string) -> list:
        with pd.ExcelFile(string) as xls:
            all_sheets = xls.sheet_names
        return all_sheets
    

if __name__ == "__main__":
    MyExcel = ExcelIO(r"Excel\chinese.xlsx", "Sheet1", "F", "B", 3)
    MyExcel.delete_score(5)