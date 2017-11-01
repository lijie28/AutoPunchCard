# -*- coding:utf-8 -*-
import xlrd
import xlutils

from xlutils.copy import copy
import time
import re
# import sys
# reload(sys)
# sys.setdefaultencoding('utf-8')


# my_name = '谢克械'.decode("utf-8")
my_id = '19388'.decode("utf-8")
my_info = []
my_data = ['占位id','占位名字','占位日期','否','否','占位上下午','个人原因','忘记打卡','无','占位上下午','物联天下']

def setExcel(data_excel_name,copy_name,final_name):
    #获取数据的excel
    data = xlrd.open_workbook(data_excel_name)
    # 获取一个工作表

    table = data.sheets()[0]  # 通过索引顺序获取
    # table = data.sheet_by_index(0)  # 通过索引顺序获取
    # table = data.sheet_by_name(u'创意车街')  # 通过名称获取

    print '表', table.nrows, table.ncols

    #生成的excel
    copy_data = xlrd.open_workbook(copy_name, formatting_info=True)
    new_row = copy_data.sheets()[0].nrows
    new_excel = copy(copy_data) # 用xlutils提供的copy方法将xlrd的对象转化为xlwt的对象
    new_table = new_excel.get_sheet(0) # 用xlwt对象的方法获得要操作的sheet

    # print table.col_values(3)
    unpunch_card = []
    for i in range(len(table.col_values(3))):
         # print name,'unicode:',u'%s'%name
        name = table.col_values(3)[i]
        id = table.col_values(2)[i]
        if my_id == id:
            record = table.col_values(44)[i]
            my_info.append(name)
            slist = record.split('>')
            # print "slict,",slist
            for strI in slist:
                if strI != '':
                    string = re.sub("[<>]", "", strI)  # 去尖括号
                    english_only = ''.join(x for x in string if ord(x) < 256)  # 去中文
                    
                    print '未打卡原记录:',u'%s'%string

                    # 转换成时间数组
                    timeArray = time.strptime(english_only, "%m-%d %H:%M")
                    tm_hour = timeArray[3]  # 具体哪个小时
                    str_time = ''
                    if 7 < tm_hour < 10:
                        str_time = '上午上班'
                    elif 11 < tm_hour <= 12:
                        str_time = '上午下班'
                    elif 12 < tm_hour < 15:
                        str_time = '下午上班'
                    elif 15 < tm_hour < 19:
                        str_time = '下午下班'
                    else:
                        str_time = '未知时间'


                    # print '时间是：',str_time
                    dt_new = time.strftime("2017/%m/%d", timeArray)
                    print '(', i,')', string,'具体时间：', english_only, '转化：', dt_new, str_time

                    #excel 写数据0
                    setPunchDatas(new_table,new_row,my_id,name,dt_new,str_time.decode("utf-8"))
                    new_row = new_row+1

                    unpunch_card.append(english_only)

    print '我的记录:', "有", len(my_info), '个', '\n', '未打卡有：', unpunch_card, len(unpunch_card), '个'
    new_excel.save(final_name) # xlwt对象的保存方法，这时便覆盖掉了原来的excel



def setPunchDatas(table,row,id,name,date,noon):
    # 获取一个工作表

    # table = toset_data.sheets()[0]  # 通过索引顺序获取
    # table = data.sheet_by_index(0)  # 通过索引顺序获取
    # table = data.sheet_by_name(u'创意车街')  # 通过名称获取
    # rows = table.nrows
    for index in range(len(my_data)):
        table.write(row, index, my_data[index].decode("utf-8"))
    table.write(row, 0, id)
    table.write(row, 1, name)
    table.write(row, 2, date)
    table.write(row, 5, noon)
    table.write(row, 9, noon)

    # excel.save(save_name) # xlwt对象的保存方法，这时便覆盖掉了原来的excel

    # print '表', table.nrows, table.ncols

    # table.cell(int(table.nrows)+1, 1).value = 'test'

setExcel('创意车街10月考勤.xlsx','test.xls','output.xls')
# setPunchDatas('test.xls','testsave.xls')

# s = "中文bab#$%$#%#$"
# r = re.sub("[A-Za-z0-9\[\`\~\!\@\#\$\^\&\*\(\)\=\|\{\}\'\:\;\'\,\[\]\.\<\>\/\?\~\！\@\#\\\&\*\%]", "", s)
# r = re.sub("[\u4e00-\u9fa5]", "", s)
# english_only = ''.join(char for char in s if not '\u4e00' <= char <= '\u9fff')

# print english_only

# 获取整行和整列的值（数组
# table.row_values(1)

# table.col_values(1)

# # 获取行数和列数
# nrows = table.nrows

# ncols = table.ncols

# # 循环行列表数据
# for i in range(nrows):
#     print table.row_values(i)

# # 单元格
# cell_A1 = table.cell(0, 0).value

# cell_C4 = table.cell(2, 3).value

# # 使用行列索引
# cell_A1 = table.row(0)[0].value

# cell_A2 = table.col(1)[0].value

# # 简单的写入
# row = 0

# col = 0

# # 类型 0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
# ctype = 1
# value = '单元格的值'

# xf = 0  # 扩展的格式化

# table.put_cell(row, col, ctype, value, xf)

# table.cell(0, 0)  # 单元格的值'

# table.cell(0, 0).value  # 单元格的值'
