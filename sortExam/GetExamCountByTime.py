# coding=utf-8
import sys, os
# 解析 xls
import xlrd
# 解析 xlsx
import openpyxl
import math
from tools import GetExcelDatas

# ///////////////////////////////////////////////////
# 获取每个时间段的考试科目人数
# ///////////////////////////////////////////////////

def getTime(row):
    return row.getValue(14)

def getExamNo(row):
    return row.getValue(15)

def getExamName(row):
    return row.getValue(16)

def getPlace(row):
    return row.getValue(18)

class Exam:
    def __init__(self):
        self.no = None
        self.name = None
        self.place = None
        self.count = 0

class rObj:
    def __init__(self):
        self.time = None
        self.exams = []

def findObj(arr, time):
    for obj in arr:
        if obj.time == time:
            return obj
    return None

g_sort = [
    [22106,11293,42768,22417,22045,22502,22610,22617,22219,22542],
    [22130,11009,22508,22109,22608,22208,11568,11171,11288,22224,11164,11568,11344,11333,24022,11110,11073],
    [22099,11291,42776,22027,22322,22625,23963,22412,22218,11141],
    [11129,11067,22403,11253,22098,22247,42732,42717,22019,11340,22623,11542,22225,22528],
    [22136,23950,22097,11123,22251,22175,11021,11069,22505,24120,42721,11313,22437,11334,22129,22626,22228,11054,11192,11251,22202],
    [24156,11575,42723,42790,24154,22180,11439,11257,22233,22517,11258,22110,22227,22332,22332],
    [11080,22094,24010,22238,22196,22108,11621,11308,22624,22226,42772,22509,42722],
    [11620,22246,11108,11254,22107,22114,22668,22072,22223,11250,22320,44988,11289,22047,23998,24046,42748,22511,24153]
]

def getExamsContent(time, exams, index):
    global g_sort
    sort = g_sort[index]
    done = []
    content = ""
    for no in sort:
        for e in exams:
            if e.no == str(no):
                content += time + "," + e.no + "," + e.name + "," + str(e.count) + "," + e.place + "\n"
                done.append(str(no))
                break
    for e in exams:
        if e.no not in done:
            content += time + "," + e.no + "," + e.name + "," + str(e.count) + "," + e.place + "\n"
    return content
    
if __name__ == '__main__':
    print("__main__>>>>>>>>>>>>>>>>>>>>>>>>")
    haveHead = True
    rowArr = GetExcelDatas.getExcelData("简阳纸考23春(1).xlsx", haveHead)
    print(len(rowArr))
    
    index = 0
    result = []
    for row in rowArr:
        if haveHead and index == 0:
            index += 1
            continue
        time = getTime(row)
        o = findObj(result, time)
        if o is None:
            o = rObj()
            result.append(o)
            print("append------------------")
            o.time = time
        
        exam = None
        no = getExamNo(row)
        for e in o.exams:
            if e.no == no:
                exam = e
                break
        if exam is None:
            exam = Exam()
            o.exams.append(exam)
            exam.no = no
            exam.name = getExamName(row)
            exam.place = getPlace(row)
        exam.count += 1
        index += 1
        
    print("rrrrrrrrrrrrrrrrr", len(result))
    content = ""
    index = 0
    count = 0
    for r in result:
        time = r.time
        count += len(r.exams)
        content += getExamsContent(time, r.exams, index)
        index += 1
    with open("resultx.csv", "w") as f:
        f.write(content)        
    print("end=========================", count)
    