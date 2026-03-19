import json

def file_loader(filepath, ENCODING = 'utf-8'):
    """
    file_loader 的 解释
    
    :param filepath: 文件路径
    :param ENCODING: 编码格式

    用来加载备份文件夹中的姓名数据
    """
    with open(filepath, encoding=ENCODING) as fp:
        return json.load(fp)
    
def setting_loader(filepath, ENCODING = 'utf-8'):
    """
    setting_loader 的 Docstring
    
    :param filepath: 说明
    :param ENCODING: 说明

    用来加载备份文件夹中的设置
    """
    with open(filepath, encoding=ENCODING) as fp:
        return json.load(fp)

def file_store(filepath, files,ENCODING = 'utf-8'):
    """
    name_store 的 Docstring
    
    :param filepath: 说明
    :param files: 说明
    :param ENCODING: 说明

    用来将新加载的姓名单重新备份
    """
    with open(filepath, mode = 'w', encoding=ENCODING) as fp:
        json.dump(files, fp)

