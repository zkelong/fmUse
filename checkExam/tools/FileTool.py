import csv
import json
import os

import chardet
from hashlib import md5


# 用默认编辑器打开指定文件
def editFile(file_path):
    if os.path.isfile(file_path):
        os.startfile(file_path)
    else:
        return "打开失败，未找到文件：{}".format(file_path)

# ################### 文件属性 ###################
# 从路径中获取文件名
def getFileNameFromPath(path):
    return os.path.basename(path)


# 获取文件名和扩展名
def getFileNameAndExtension(path):
    filename = getFileNameFromPath(path)
    info = os.path.splitext(filename)
    return info[0], info[1]


# 获取文件编码
def getFileEncoding(file):
    with open(file, 'rb') as f:
        result = chardet.detect(f.read())
    return result['encoding']


# ################### 操作文件 ###################
# 文件重命名 newName:新名字，newExtension:新扩展名
def renameFile(file, newName=None, newExtension=None):
    oldName, oldExtension = getFileNameAndExtension(file)
    oldFileName = oldName + oldExtension
    path = file.rstrip(oldFileName)
    if newName:
        newFileName = newName
    else:
        newFileName = oldName
    if newExtension:
        newFileName += newExtension
    else:
        newFileName += oldExtension
    os.rename(file, os.path.join(path, newFileName))


# 修改目录/文件时间
def modifyDirFileTime(filePath, mTime):
    os.utime(filePath, (mTime, mTime))

# 获取文件 md5
def getMd5(filePath):
    m = md5()
    with open(filePath, "rb") as f:
        m.update(f.read())
    return m.hexdigest()


# ################### 读取文件 ###################
encodingTypes = ["UTF-8", "Unicode", "GBK", "ISO-8859-1", "UTF-16"]


# 读取 json 文件内容
def readJsonFile(file, encoding='utf-8'):
    error = None
    data = None
    if os.path.isfile(file):
        with open(file, 'r', encoding=encoding) as f:
            content = f.read()
            try:
                # 区别 json.load 和 json.loads
                data = json.loads(content)
            except json.JSONDecodeError as e:
                error = "{}语法错误：{}".format(file, e.msg)
    else:
        error = "{}文件不存在".format(file)
    return data, error


# 读取 csv 内容，尝试用不同编码方式解析
def readCsvFile(file, usedEncodings=None):
    encoding = None
    if usedEncodings is None:
        encoding = getFileEncoding(file)
    else:
        for en in encodingTypes:
            if en not in usedEncodings:
                usedEncodings.append(en)
                encoding = en
                break
        if encoding is None:
            return None, "无法解析编码:{}".format(file)
    try:
        csv_reader = csv.reader(open(file, 'r', encoding=encoding))
        rows = []
        for row in csv_reader:
            rows.append(row)
        return rows, None
    except Exception as e:
        return readCsvFile(file, usedEncodings or [])




