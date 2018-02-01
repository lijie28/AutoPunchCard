import xlrd
import xlutils
from xlutils.copy import copy
import time
import re

def creatExcel(final_name,unpunch):
    #生成的excel
    copy_data = xlrd.open_workbook('./source/output_sample.xls', formatting_info=True)
    new_row = copy_data.sheets()[0].nrows
    new_excel = copy(copy_data) # 用xlutils提供的copy方法将xlrd的对象转化为xlwt的对象
    new_table = new_excel.get_sheet(0) # 用xlwt对象的方法获得要操作的sheet

    #excel 写数据0
    for x in unpunch:
        setPunchDatas(new_table,new_row,x['id'],x['name'],x['date'],x['time'])
        new_row = new_row+1

    new_excel.save(final_name) # xlwt对象的保存方法，这时便覆盖掉了原来的excel

def checkExcel(data_excel_name,my_id):
    #获取数据的excel
    data = xlrd.open_workbook(data_excel_name)
    # 获取一个工作表
    table = data.sheets()[0]  # 通过索引顺序获取
    # print ('表', table.nrows, table.ncols)
    # print table.col_values(3)
    unpunch = []
    showmsglist = []
    my_info = []
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

                    dt_new = time.strftime("2018/%m/%d", timeArray)
                    showmsg =  ('(%d)' %i) + name + string +',具体时间：' + english_only +',=====>>'+ dt_new + ' ' + str_time
                    unpunch_detail = {
                        'id':my_id,
                        'name':name,
                        'date':dt_new,
                        'time':str_time
                    }
                    # print (showmsg)
                    showmsglist.append(showmsg)
                    unpunch.append(unpunch_detail)


    # print ('我的记录:', "有", len(my_info), '个', '\n', '未打卡有：', unpunch, len(unpunch), '个')
    return (showmsglist,unpunch)


def setPunchDatas(table,row,id,name,date,noon):
    my_data = ['占位id','占位名字','占位日期','否','否','占位上下午','个人原因','忘记打卡','无','占位上下午','物联天下']
    # 获取一个工作表
    for index in range(len(my_data)):
        table.write(row, index, my_data[index])
    table.write(row, 0, id)
    table.write(row, 1, name)
    table.write(row, 2, date)
    table.write(row, 5, noon)
    table.write(row, 9, noon)
    
if __name__ == "__main__":
    result = checkExcel('/Users/lijie/Documents/GitHub/AutoPunchCard/创意车街考勤.xlsx','19489')
    print (result[1])
    creatExcel('tehah.xls',result[1])