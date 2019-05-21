import xlrd
import xlwt
from xlutils.copy import copy

def write_excel_xls(path, sheet_name, value):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlwt.Workbook()  # 新建一个工作簿
    sheet = workbook.add_sheet(sheet_name)  # 在工作簿中新建一个表格
    for i in range(0, index):
        for j in range(0, len(value[i])):
            sheet.write(i, j, value[i][j])  # 像表格中写入数据（对应的行和列）
    workbook.save(path)  # 保存工作簿
    print("xls格式表格写入数据成功！")

def read_excel_xls(path):
    data = []
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    if worksheet.nrows == 1:
        print("目前是第一行")
    else:
        for i in range(1, worksheet.nrows): #从第二行取值
            dataTemp = []
            for j in range(0, worksheet.ncols):
                #print(worksheet.cell_value(i, j), "\t", end="")  # 逐行逐列读取数据
                dataTemp.append(worksheet.cell_value(i, j))
            data.append(dataTemp)
    return data
     
def write_excel_xls_append_norepeat(path, value):
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
    rid = 0
    for i in range(0, len(value)):
        data = read_excel_xls(path)
        data_temp = []
        for m in range(0,len(data)):
            data_temp.append(data[m][1:len(data[m])])
        value_temp = []
        for m in range(0,len(value)):
            value_temp.append(value[m][1:len(value[m])])
        
        if value_temp[i] not in data_temp:
            for j in range(0, len(value[i])):
                new_worksheet.write(rid+rows_old, j, value[i][j])  # 追加写入数据，注意是从i+rows_old行开始写入
            rid = rid + 1
            new_workbook.save(path)  # 保存工作簿
            print("xls格式表格【追加】写入数据成功！")
        else:
            print("数据重复")
