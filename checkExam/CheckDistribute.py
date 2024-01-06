# coding=utf-8
from tools import GetExcelDatas
from tools import FileTool
from CheckExamPlan import CheckAllExamInfo
# 解析 xls
import xlrd
# 解析 xlsx
import openpyxl

# ///////////////////////////////////////////////////
# 检查回收透视表是否正确
# ///////////////////////////////////////////////////

# 考试场次信息
class ExamInfo:
    def __init__(self, row):
        # 试卷号,试卷名称,试卷份数,考场号 在表中的第几列，从0开始
        self.infoIndex = [0, 1, 2, 3]
        # 试卷号
        self.paperNo = str(row.getValue(self.infoIndex[0])).strip()
        # 试卷名
        self.paperName = str(row.getValue(self.infoIndex[1])).strip()
        # 份数/人数
        self.num = row.getValue(self.infoIndex[2])
        # 考场号
        self.placeNo = str(row.getValue(self.infoIndex[3])).strip()

    def getHeadValues(self):
        return ["试卷号","考试科目","份数","考场号"]
    
    def getValues(self):
        return [self.paperNo, self.paperName, self.num, self.placeNo]

# 所有回收透视
class CheckRecoveryInfo:
    def __init__(self, file=None):
        self.file = file or "23秋金牛分时间安排表.xlsx"
        # 标准信息（时间安排表的信息）
        self.baseData = None
        self.sheetIndex = 4
        self.sheetTitle = None
        self.infos = []
        self.parseInfos()

    def parseInfos(self):
        rowArr, self.sheetTitle = GetExcelDatas.getExcelData(self.file, True, self.sheetIndex)
        for row in rowArr:
            if not row.isHead:
                newInfo = ExamInfo(row)
                self.checkRepeat(newInfo)
                self.infos.append(newInfo)

    # 是否有重复的信息：卷号 相同
    def checkRepeat(self, newInfo):
        for info in self.infos:
            if info.paperNo == newInfo.paperNo:
                print("重复信息：试卷号（{}）".format(info.paperNo))

    # 获取指定试卷号的信息
    def findInfoByPaperNo(self, paperNo):
        for info in self.infos:
            if info.paperNo == paperNo:
                return info

    def getBaseDate(self):
        if self.baseData is None:
            self.baseData = CheckAllExamInfo()

    # 检查数据是否正确
    def checkInfos(self):
        print("\n===============开始检查{}".format(self.sheetTitle))
        self.getBaseDate()
        # 当前表有的，而标准信息缺失的
        missInfo = []
        missInfoStr = ""
        for info in self.infos:
            baseInfo = self.baseData.findInfoByPaperNo(info.paperNo)
            if baseInfo is None:
                missInfo.append(info)
                missInfoStr += "试卷：{}\n".format(info.paperNo)
            else:
                self.compareInfo(baseInfo, info)
        if len(missInfoStr):
            print("标准信息缺失：\n" + missInfoStr)

        moreInfo = []
        moreInfoStr = ""
        for info in self.baseData.infos:
            sInfo = self.findInfoByPaperNo(info.paperNo)
            if sInfo is None:
                moreInfo.append(info)
                moreInfoStr += "试卷：{}\n".format(info.paperNo)
        if len(moreInfoStr):
            print("标准信息多了：\n" + missInfoStr)
        

    # 对比数据
    def compareInfo(self, baseInfo, info):
        diff = ""
        # 试卷名
        if baseInfo.paperName != info.paperName:
            diff += "试卷名：基础({}),对比({});".format(baseInfo.paperName, info.paperName)
        if str(baseInfo.num) != str(info.num):
            diff += "人数：基础({}),对比({});".format(baseInfo.num, info.num)
        if baseInfo.placeNo != info.placeNo:
            diff += "袋数：基础({}),对比({});".format(baseInfo.placeNo, info.placeNo)
        if len(diff) > 0:
            print("考试{}差异：{}".format(info.paperNo, diff))

    # 考试信息写入文件
    def writeToFile(self):
        if len(self.infos) < 1:
            print("没有数据,无法写入")
            return
            
        self.getBaseDate()
        nFile = self.sheetTitle + "_result.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        # 从第一行开始
        excelRowIndex = 1
        self.writeExcelRowInfo(sheet, excelRowIndex, self.infos[0].getHeadValues() + self.infos[0].getHeadValues())
        excelRowIndex += 1
        for info in self.infos:
            baseInfo = self.baseData.findInfoByPaperNo(info.paperNo)
            baseValues = []
            if baseInfo is None:
                baseValues = ["", "", "", ""]
            else:
                baseValues = [baseInfo.paperNo, baseInfo.paperName, baseInfo.num, baseInfo.placeNo]
            self.writeExcelRowInfo(sheet, excelRowIndex, info.getValues() + baseValues)
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

# do
checkInfo = CheckRecoveryInfo()
checkInfo.checkInfos()
checkInfo.writeToFile()