# coding=utf-8
from tools import GetExcelDatas
from tools import FileTool
# 解析 xls
import xlrd
# 解析 xlsx
import openpyxl

# ///////////////////////////////////////////////////
# 根据名单表整理出每场考试的人数等信息
# ///////////////////////////////////////////////////

# 考试场次信息
class ExamInfo:
    def __init__(self, row):
        # 试卷号,试卷名称,考场号,教室,日期 在表中的第几列，从0开始
        self.infoIndex = [13, 14, 6, 19, 11]
        # 试卷号
        self.paperNo = row.getValue(self.infoIndex[0]).strip()
        # 试卷名
        self.paperName = row.getValue(self.infoIndex[1]).strip()
        # 考场号
        self.placeNo = row.getValue(self.infoIndex[2]).strip()
        # 教室
        self.placeName = row.getValue(self.infoIndex[3]).strip()
        # 日期
        self.date = row.getValue(self.infoIndex[4]).strip()
        # 人数
        self.num = 1

    def getCompareDate(self):
        return self.date.replace(" ", "")[0:15]
    
    def addNew(self, info):
        diff = ""
        if self.paperName != info.paperName:
            diff += "试卷名：{},{}".format(self.paperName, info.paperName)
        if self.placeNo != info.placeNo:
            diff += "考场号：{},{}".format(self.placeNo, info.placeNo)
        if self.placeName != info.placeName:
            diff += "教室：{},{}".format(self.placeName, info.placeName)
        if self.date != info.date:
            diff += "日期：{},{}".format(self.date, info.date)
        if len(diff) > 0:
            print("试卷号{}信息不同：{}".format(self.paperNo, diff))
        self.num += 1

    def getHeadValues(self):
        return ["试卷号","试卷名称","考场号","教室","日期","人数"]
    
    def getValues(self):
        return [self.paperNo, self.paperName, self.placeNo, self.placeName, self.date, self.num]

# 所有考试信息
class AllExamInfo:
    def __init__(self, file=None):
        self.file = file or "简阳考试名单.xlsx"
        self.sheetTitle = None
        self.infos = []
        self.parseInfos()

    def parseInfos(self):
        rowArr, self.sheetTitle  = GetExcelDatas.getExcelData(self.file, True)
        for row in rowArr:
            if not row.isHead:
                newInfo = ExamInfo(row)
                info = self.findInfoByPaperNoPlaceNo(newInfo.paperNo, newInfo.placeNo)
                if info:
                    info.addNew(newInfo)
                else:
                    self.infos.append(newInfo)

    # 获取指定试卷号-考场号的信息(同一试卷可分两场)
    def findInfoByPaperNoPlaceNo(self, paperNo, placeNo):
        for info in self.infos:
            if info.paperNo == paperNo and info.placeNo == placeNo:
                return info

    # 考试信息写入文件
    def writeToFile(self):
        if len(self.infos) < 1:
            print("没有数据,无法写入")
            return
        oName, oExtension = FileTool.getFileNameAndExtension(self.file)
        nFile = oName + "_result" + oExtension
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        # 从第一行开始
        excelRowIndex = 1
        self.writeExcelRowInfo(sheet, excelRowIndex, self.infos[0].getHeadValues())
        excelRowIndex += 1
        for info in self.infos:
            self.writeExcelRowInfo(sheet, excelRowIndex, info.getValues())
            excelRowIndex += 1
        workbook.save(nFile)
        workbook.close()
        print("导出完成：{}".format(nFile))

    def writeExcelRowInfo(self, sheet, rowIndex, values):
        # row, column 从 1 开始
        for _index, _value in enumerate(values):
            cell = sheet.cell(row=rowIndex, column=_index + 1)
            # 科学计数法问题，设置单元格格式
            cell.number_format = '0'  # 或者使用 '0.00' 等形式，确保数字以常规格式显示，而非科学计数法
            cell.value = _value

# test
# allInfo = AllExamInfo()
# allInfo.writeToFile()