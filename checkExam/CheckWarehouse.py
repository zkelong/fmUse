# coding=utf-8
from tools import GetExcelDatas
from tools import FileTool
from CheckExamPlan import CheckAllExamInfo
# 解析 xls
import xlrd
# 解析 xlsx
import openpyxl

# ///////////////////////////////////////////////////
# 检查入库单是否正确
# ///////////////////////////////////////////////////

# 考试场次信息
class ExamInfo:
    def __init__(self, row):
        # 试卷号,试卷名称,考场人数,31份装（袋）,26份装（袋）,11份装（袋）,5份装（袋),合计（袋）,时间
        self.infoIndex = [0, 1, 2, 3, 4, 5, 6, 7, 8]
        # 试卷号
        self.paperNo = str(row.getValue(self.infoIndex[0])).strip()
        # 试卷名
        self.paperName = row.getValue(self.infoIndex[1]).strip()
        # 份数/人数
        self.num = int(row.getValue(self.infoIndex[2]))
        # 31份装（袋）
        self.pack31 = int(row.getValue(self.infoIndex[3]))
        # 26份装（袋）
        self.pack26 = int(row.getValue(self.infoIndex[4]))
        # 11份装（袋）
        self.pack11 = int(row.getValue(self.infoIndex[5]))
        # 5份装（袋）
        self.pack5 = int(row.getValue(self.infoIndex[6]))
        # 合计（袋）
        self.packNum = int(row.getValue(self.infoIndex[7]))
        # 时间
        self.date = row.getValue(self.infoIndex[8])

    def getCompareDate(self):
        return self.date.replace(" ", "")[0:15]

    def getHeadValues(self):
        return ["试卷号","试卷名称","考场人数","31份装（袋）","26份装（袋）","11份装（袋）","5份装（袋）","合计（袋）","时间"]
    
    def getValues(self):
        return [self.paperNo, self.paperName, self.num, self.pack31, self.pack26, self.pack11, self.pack5, self.packNum, self.date]

# 所有入库信息
class CheckWareHouseInfo:
    def __init__(self, file=None):
        self.file = file or "23秋金牛分时间安排表.xlsx"
        # 标准信息（时间安排表的信息）
        self.baseData = None
        self.sheetIndex = 2
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

    # 是否有重复的信息：试卷号 相同
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

        paperNum = info.num
        paperNum, nDiff = self.checkPack(info.pack31, paperNum, 31, 26)
        if len(nDiff) > 0:
            diff += nDiff
        # paperNum, nDiff = self.checkPack(info.pack26, paperNum, 26, 11)
        # if len(nDiff) > 0:
        #     diff += nDiff
        paperNum, nDiff = self.checkPack(info.pack11, paperNum, 11, 5)
        if len(nDiff) > 0:
            diff += nDiff
        paperNum, nDiff = self.checkPack(info.pack5, paperNum, 5, 0)
        if len(nDiff) > 0:
            diff += nDiff

        allPackNum = info.pack31 + info.pack26 + info.pack11 + info.pack5
        if info.packNum != allPackNum:
            diff += "袋数：目前({}),应该({});".format(info.packNum, allPackNum)
        if baseInfo.getCompareDate() != info.getCompareDate():
            diff += "日期：基础({}),对比({})".format(baseInfo.date, info.date)
        if len(diff) > 0:
            print("考试{}差异：{}".format(info.paperNo, diff))

    def checkPack(self, nowPackNum, paperNum, nowPackPaperNum, nextPackPaperNum):
        diff = ""
        if paperNum < 1:
            return 0, diff
        if paperNum > nowPackNum:
            # 超过当前袋包含的数量
            packNum = paperNum // nowPackPaperNum
            # 剩余数量比下一袋能包含的多
            leftNum = paperNum % nowPackPaperNum
            if leftNum > nextPackPaperNum:
                packNum += 1
                leftNum = 0
            if nowPackNum != packNum:
                diff += "{}份装（袋）数量错误，应该({}),实际({});".format(nowPackPaperNum, packNum, nowPackNum)
            return leftNum, diff
        return paperNum, diff

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
        self.writeExcelRowInfo(sheet, excelRowIndex, self.infos[0].getHeadValues() + ["试卷号","试卷名","人数","时间"])
        excelRowIndex += 1
        for info in self.infos:
            baseInfo = self.baseData.findInfoByPaperNo(info.paperNo)
            baseValues = []
            if baseInfo is None:
                baseValues = ["", "", "", ""]
            else:
                baseValues = [baseInfo.paperNo, baseInfo.paperName, baseInfo.num, baseInfo.date]
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
checkInfo = CheckWareHouseInfo()
checkInfo.checkInfos()
checkInfo.writeToFile()