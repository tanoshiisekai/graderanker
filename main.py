from pyexcel_xls import get_data as get_data_xls
from pyexcel_xlsx import get_data as get_data_xlsx
import xlsxwriter
import os


dirname = "C:/datacompare/"

sourcefilename = "test.xlsx"
destfilename = "test1.xlsx"

sourcetablename = "Sheet1"
desttablename = "Sheet1"

idcol = 'A'
namecol = 'B'
lastcol1 = 'C'
gradecols = ['D', 'E', 'F', 'G', 'H', 'I', 'J']

rowstart1 = 1
rowcount = 43
rowstart2 = 1

if os.path.exists(dirname + destfilename):
    os.remove(dirname + destfilename)

workbook = xlsxwriter.Workbook(dirname + destfilename)
worksheet1 = workbook.add_worksheet()

formatmax = workbook.add_format({'bg_color': '#6495ED',
                                 'font_color': '#FFFFFF'})  # 单科第一

formatfailed = workbook.add_format({'bg_color': '#E9F01D',
                                    'font_color': '#F01B2D'})  # 不及格

formatbetter = workbook.add_format({'bg_color': '#56A36C',
                                    'font_color': '#FFFFFF'})  # 进步明显

formatworse = workbook.add_format({'bg_color': '#F01B2D',
                                   'font_color': '#FFFFFF'})  # 退步明显

formatbottom10 = workbook.add_format({'bg_color': '#FF84BA',
                                    'font_color': '#FFFFFF'})  # 总分后10名

formattop10 = workbook.add_format({'bg_color': '#0079BA',
                                    'font_color': '#FFFFFF'})  # 总分前10名


def userfunction(datainmem):
    """
    用户自定义操作，例如排序等
    注意不要修改列的顺序
    :param datainmem: 计算排名后的数据
    :return:
    """
    titlelist = [x[0] for x in datainmem]
    datainmem = [x[1:] for x in datainmem]
    datainmem = list(zip(*datainmem))
    datainmem.sort(key=lambda x: (x[15], x[1]), reverse=True)
    datainmem.insert(0, titlelist)
    return datainmem


def getcolnum(colname):
    """
    列名转下标，0起
    :param colname:列名
    :return:下标
    """
    thesum = 0
    length = len(colname)
    loop = length - 1
    while loop >= 0:
        thesum = thesum + (ord(colname[length - loop - 1]) - ord('A') + 1) * (26 ** loop)
        loop = loop - 1
    return thesum - 1


def colnumgenerator():
    sourcevalue = 1
    while True:
        valuestr = ""
        remainlist = []
        value = sourcevalue
        while value:
            remain = value % 26
            value = value // 26
            if remain == 0:
                remainlist.append(26)
                value = value - 1
            else:
                remainlist.append(remain)
        remainlist.reverse()
        for rem in remainlist:
            valuestr = valuestr + chr(ord('A') + rem - 1)
        sourcevalue = sourcevalue + 1
        yield valuestr


def getcolname(colnum):
    """
    下标转列名，0起
    :param colnum:列标
    :return:
    """
    count = 0
    for i in colnumgenerator():
        count = count + 1
        if count == colnum + 1:
            return i


def readdata():
    """
    读取源数据
    :return: 数据对象
    """
    if sourcefilename.endswith(".xls"):
        f1 = get_data_xls(dirname + sourcefilename)[sourcetablename][rowstart1 - 1:rowstart1 + rowcount]
    if sourcefilename.endswith(".xlsx"):
        f1 = get_data_xlsx(dirname + sourcefilename)[sourcetablename][rowstart1 - 1:rowstart1 + rowcount]
    return f1


def getrank(coldatalist):
    """
    得到中国式排名
    :param coldatalist: （序号，分数，排名）三元组列表
    :return:
    """
    templist = []
    for i in range(1, len(coldatalist) + 1):
        if coldatalist[i - 1] is None:
            return None
        templist.append([i, coldatalist[i - 1], 0])
    templist.sort(key=lambda x: x[1], reverse=True)
    p = 0
    r = 1
    t = 1
    templist[p][2] = r
    while True:
        p = p + 1
        if p == len(templist):
            break
        t = t + 1
        if templist[p][1] != templist[p - 1][1]:
            r = r + 1
            flagr = 0
        else:
            flagr = 1
        if flagr == 0:
            templist[p][2] = t
            r = t
        else:
            templist[p][2] = r
    templist.sort(key=lambda x: x[0])
    return templist


def sumwithnone(numlist):
    numlist = [x for x in numlist if x]
    return sum(numlist)


if __name__ == "__main__":
    # 读取数据
    data = readdata()
    datainmemory = []  # 在内存中组织数据
    subcols = []  # 记录科目列
    sumlist = []  # 记录总分列
    # 写入序号
    idlist = [x[getcolnum(idcol)] for x in data]
    datainmemory.append(idlist)
    # 写入姓名
    namelist =[x[getcolnum(namecol)] for x in data]
    datainmemory.append(namelist)
    # 写入上一次排名
    lastrank = [x[getcolnum(lastcol1)] for x in data]
    datainmemory.append(lastrank)
    # 写入各科成绩并标明单科第一和不及格
    startcol = getcolnum(gradecols[0])
    for col in gradecols:
        subcols.append(startcol)
        values = [x[getcolnum(col)] for x in data]
        if None in values:
            # 成绩还没有公布
            continue
        datainmemory.append(values)
        sumlist.append(values[1:])
        rankvalues = [x[2] for x in getrank(values[1:])]
        rankvalues.insert(0, "排名")
        datainmemory.append(rankvalues)
        startcol = startcol + 2
    # 写入总分及排名
    totalcol = startcol
    sumlist = [sumwithnone(x) for x in zip(*sumlist)]
    sumlist.insert(0, "总分")
    datainmemory.append(sumlist)
    startcol = startcol + 1
    rankcol = startcol
    ranklist = [x[2] for x in getrank(sumlist[1:])]
    ranklist.insert(0, "排名")
    datainmemory.append(ranklist)
    startcol = startcol + 1
    # 写入进步情况
    stepcol = startcol
    lastr = lastrank[1:]
    nowr = ranklist[1:]
    steplist = [x[0] - x[1] for x in zip(lastr, nowr)]
    steplist.insert(0, "进步")
    datainmemory.append(steplist)
    startcol = startcol + 1
    # 再写入一次姓名列，方便对齐
    datainmemory.append(namelist)
    startcol = startcol + 1
    datainmemory = userfunction(datainmemory)

    # 完成写入文件
    for i in range(1, len(datainmemory) + 1):
        worksheet1.write_row('A' + str(i), datainmemory[i - 1])

    # 添加各种标注
    for sc in subcols:
        worksheet1.conditional_format(getcolname(sc) + str(rowstart2 + 1) + ":" +
                                      getcolname(sc) + str(rowstart2 + rowcount),
                                      {
                                          'type': 'cell',
                                          'criteria': '<',
                                          'value': 60,
                                          'format': formatfailed
                                      })
        worksheet1.conditional_format(getcolname(sc) + str(rowstart2 + 1) + ":" +
                                      getcolname(sc) + str(rowstart2 + rowcount),
                                      {
                                          'type': 'top',
                                          'value': 1,
                                          'format': formatmax
                                      })
    worksheet1.conditional_format(getcolname(stepcol) + str(rowstart2 + 1) + ":" +
                                  getcolname(stepcol) + str(rowstart2 + rowcount),
                                  {
                                      'type': 'cell',
                                      'criteria': '<',
                                      'value': -5,
                                      'format': formatworse
                                  })
    worksheet1.conditional_format(getcolname(stepcol) + str(rowstart2 + 1) + ":" +
                                  getcolname(stepcol) + str(rowstart2 + rowcount),
                                  {
                                      'type': 'cell',
                                      'criteria': '>',
                                      'value': 5,
                                      'format': formatbetter
                                  })
    worksheet1.conditional_format(getcolname(rankcol) + str(rowstart2 + 1) + ":" +
                                  getcolname(rankcol) + str(rowstart2 + rowcount),
                                  {
                                      'type': 'top',
                                      'value': 10,
                                      'format': formatbottom10
                                  })
    worksheet1.conditional_format(getcolname(rankcol) + str(rowstart2 + 1) + ":" +
                                  getcolname(rankcol) + str(rowstart2 + rowcount),
                                  {
                                      'type': 'bottom',
                                      'value': 10,
                                      'format': formattop10
                                  })
    workbook.close()
