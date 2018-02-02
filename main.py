from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import sys,os,json
from punchcard import checkExcel ,creatExcel
import sip,time


class MyLineEdit(QLineEdit):
    def __init__(self,parent=None):
        super(QLineEdit,self).__init__(parent)
        ###初始化打开接受拖拽使能
        self.setAcceptDrops(True)

    def dragEnterEvent(self,event):
        event.accept()

    def dropEvent(self, event):
        ###获取拖放过来的文件的路径
        st = str(event.mimeData().urls())
        ########st就是Qt文件的路径。我们将这个路径稍作处理便可以得到我们想要的路径了
        st = st.replace("[PyQt5.QtCore.QUrl('file://","")
        st = st.replace("'), ",",")
        st = st.replace("PyQt5.QtCore.QUrl('file://","")
        st = st.replace("')]","")
        # st = st
        self.setText(st)
        print ("drag end")



class PersonalInfo():
    """docstring for PersonalInfo"""
    def __init__(self, arg=None):
        super(PersonalInfo, self).__init__()
        # self.arg = arg
        # self.open()
    def open():
        model={} #存放读取的数据
        if os.access("./source/info.json", os.F_OK):
            with open("./source/info.json",'r',encoding='utf-8') as json_file:
                model=json.load(json_file)
            return model
        else:
            return []
            pass
            # return '不存在文件'
    
    def save(key,value):
        info = PersonalInfo.open()
        info[key] = value
        print ('save info',info)
        # with open("./source/info.json",'w',encoding='utf-8') as json_file:
            # json.dump(info,json_file,ensure_ascii=False)

    def create():
        pass


class MyGridLayout(QGridLayout):
    def __init__(self, parent=None):
        super(MyGridLayout, self).__init__(parent)

    def lazyAddWidget(widget,Class,x=0,y=0):
        if widget is None:
            widget = Class
            return addWidget(widget, x, y)
        else:
            # pass
            return widget
        
        
# def lazyAddWidget(gridLayout,widget,Class,x=0,y=0):
#     if widget is None:
#         widget = Class
#         gridLayout.addWidget(widget, x, y)
#     else:
#         pass
#     return widget

class SelectDialog(QDialog):
    def __init__(self, parent=None):
        super(SelectDialog, self).__init__(parent)
        self.info = PersonalInfo.open()
        self.path = os.getcwd()[0]
        self.initUI()
        self.setWindowTitle("批补卡")
        self.resize(340, 100)
        # print self.info
        self.record_detail = None#记录详情
        self.record_num = None#记录条数
        self.work_reason = None#工作原因
        self.personal_reason = None#个人原因
        self.personal_reason_tit = None#个人原因标题
        self.btn_output = None#导出按钮
        self.outputpath = None#导出路径
        self.buttonBox = None#确定保存
        # self.selectnum = None#选择的工作原因序号
        # self.selectnum_tit = None#选择的工作原因序号标题

    def getInfo(test,key):
        if test.info:
            if key in test.info:
                return test.info[key]
        return ''


    def initUI(self):
        worknum = self.getInfo('num')
        self.path  = self.getInfo('filepath')

        grid = QGridLayout()
        self.grd = grid
        grid.addWidget(QLabel("考勤表："), 0, 0)
        self.pathLineEdit = MyLineEdit()
        # self.pathLineEdit.setFixedWidth(400)
        self.pathLineEdit.setText(self.path)
        self.pathLineEdit.setDragEnabled = True
        self.pathLineEdit.setDropEnabled = True
        grid.addWidget(self.pathLineEdit, 0, 1)

        button = QPushButton("更改")
        button.clicked.connect(self.changePath)
        grid.addWidget(button, 0, 2)

        self.tfnum = QLineEdit()
        self.tfnum.setFixedWidth(100)
        self.tfnum.setText(worknum)
        self.tfnum.move(100,50)
        grid.addWidget(QLabel("工号："), 2, 0)
        grid.addWidget(self.tfnum, 2, 1)

        button3 = QPushButton("查看")
        button3.clicked.connect(self.check)
        grid.addWidget(button3, 2, 2)

        self.setLayout(grid)


    def check(self):#检查
        result = checkExcel(self.path,self.tfnum.text())
        self.result = result[1]
        print ('路径:',self.path,'工号：',self.tfnum.text(),'条数：',len(result[1]))

        #未打卡记录表
        if self.record_detail is None :
            self.record_detail = QListWidget()
            self.record_detail.addItems(result[0])
            self.grd.addWidget(self.record_detail, 3, 1)
        else:
            for i in range(self.record_detail.count())[::-1]:
                select_item = self.record_detail.takeItem(i)
                self.record_detail.removeItemWidget(select_item)
            self.record_detail.addItems(result[0])  
        self.record_detail.setSelectionMode(QAbstractItemView.MultiSelection)
        # self.record_detail.setCurrentRow(0) 
        # self.record_detail.setCurrentRow(1) 
        # print ('selectedItems,',self.record_detail.selectedItems().count())

        # 未打卡数
        if self.record_num is None :
            self.record_num = QLabel()
            # self.record_num.setText("有%d条未打卡记录" % len(result[1]))
            self.grd.addWidget(self.record_num, 4, 1)
        # else:
        if len(result[1]) >2:
            self.record_num.setText("有%d条未打卡记录,工作原因只能选2条" % len(result[1]))
        else:
            self.record_num.setText("有%d条未打卡记录" % len(result[1]))


        #工作原因
        if self.work_reason is None:
            self.work_reason = QLineEdit()
            self.grd.addWidget(QLabel("工作原因："), 6, 0)
            self.grd.addWidget(self.work_reason, 6, 1)
        self.work_reason.setText(self.getInfo('work_reason'))


        if len(result[1])>2:
            # print ('记录大于2')
            if self.personal_reason is None:
                self.personal_reason = QLineEdit()
                self.personal_reason_tit = QLabel("个人原因：")
                self.grd.addWidget(self.personal_reason_tit, 7, 0)
                self.grd.addWidget(self.personal_reason, 7, 1)
                self.personal_reason.setText(self.getInfo('personal_reason'))

            # if self.selectnum is None:
            #     self.selectnum = QLineEdit()
            #     self.grd.addWidget(self.selectnum, 6, 2)
            #     self.selectnum.setText("请填写工作原因序号")
                
        else:
            if self.personal_reason :
                # print ('chekck')
                self.grd.removeWidget(self.personal_reason)
                self.grd.removeWidget(self.personal_reason_tit)

                sip.delete(self.personal_reason)
                sip.delete(self.personal_reason_tit)

                self.personal_reason = None
                self.personal_reason_tit = None

            # if self.selectnum :
            #     self.grd.removeWidget(self.selectnum)
            #     sip.delete(self.selectnum)
            #     self.selectnum = None



        if self.btn_output is None:
            self.btn_output = QPushButton("设置")
            self.btn_output.clicked.connect(self.output)
            self.grd.addWidget(QLabel("导出路径："), 8, 0)   
            self.grd.addWidget(self.btn_output, 8, 2)


        if self.outputpath is None: 
            self.outputpath = QLabel()
            self.grd.addWidget(self.outputpath, 8, 1)
        self.outputpath.setText(self.getInfo('output'))

        if self.buttonBox is None:
            self.buttonBox = QPushButton("导出")
            self.buttonBox.clicked.connect(self.confirm) 
            self.grd.addWidget(self.buttonBox, 9, 1)

    def output(self):
        open = QFileDialog()
        outputpath_str = open.getExistingDirectory()
        self.outputpath.setText(outputpath_str)
        # print ('set output',self.outputpath.text())    

        # if self.buttonBox is None:
        #     self.buttonBox = QDialogButtonBox()
        #     self.buttonBox.setOrientation(Qt.Horizontal)  # 设置为水平方向
        #     self.buttonBox.setStandardButtons(QDialogButtonBox.lala|QDialogButtonBox.Cancel)
        #     self.grd.addWidget(self.buttonBox, 9, 1)
        # self.buttonBox.accepted.connect(self.test)  # 确定
        # self.buttonBox.rejected.connect(self.test)  # 取消
      
    def confirm(self):
        # print ('confirm')
        if self.work_reason.text() == '':
            QMessageBox.information(self,"","请填写工作原因") 
        elif self.outputpath.text() == '':
            QMessageBox.information(self,"","请设置导出路径")
        elif self.tfnum.text() == '':
            QMessageBox.information(self,"","请填写工号")
        elif self.personal_reason :
            if (self.personal_reason.text() == ''):
                QMessageBox.information(self,"","请填写个人原因")
            else :  
                # print ('ok',self.result)
                PersonalInfo.save('filepath',self.pathLineEdit.text())
                PersonalInfo.save('num',self.tfnum.text())
                PersonalInfo.save('work_reason',self.work_reason.text())
                PersonalInfo.save('output',self.outputpath.text())
                PersonalInfo.save('personal_reason',self.personal_reason.text())

                combox=self.record_detail.selectedItems() 
                lnum = []
                for row in range(self.record_detail.count()):
                    for x in self.record_detail.selectedItems() :
                        if self.record_detail.item(row).text() == x.text():
                            # print ('内容：',x.text(),row)
                            lnum.append(row)
                    # print ('内容：',self.record_detail.item(row).text())
                print ('内容有：',lnum)
                self.outputExcel(lnum)
        else:
            # print ('ok',self.result)
            PersonalInfo.save('filepath',self.path)
            PersonalInfo.save('num',self.tfnum.text())
            PersonalInfo.save('work_reason',self.work_reason.text())
            PersonalInfo.save('output',self.outputpath.text())
            self.outputExcel()
            # self.record_detail.currentText()
    
    def outputExcel(self,list=[0,1]):
        if len(self.result) <= 2 :
            for dic in self.result:
                dic['type'] = '工作原因'
                dic['detail'] = self.work_reason.text()

            # print ('ok',self.result)
        elif not len(list)==2:
            QMessageBox.information(self,"","请选择两个工作原因")
            return
        else:
            # print ('list有',list)
            for i in range(len(self.result)):    
                for x in list:
                    # print ('x,',x,'i,',i)
                    if i == x :
                        self.result[i]['type'] = '工作原因'
                        self.result[i]['detail'] = self.work_reason.text()
                        break
                    else:
                        self.result[i]['type'] = '个人原因'
                        self.result[i]['detail'] = self.personal_reason.text()
                        
        
        nowtime = time.strftime('%Y-%m-%d',time.localtime(time.time()))
        print ('ok',self.result,nowtime)
        pathname = self.outputpath.text() + '/%s--%s.xls'%(self.result[0]['name'],nowtime)
        strinfo = creatExcel(pathname,self.result)
        QMessageBox.information(self,"",strinfo) 
            

    def changePath(self):
        # print ('enter')
        open = QFileDialog()
        getpath=open.getOpenFileName()[0]
            
        if (getpath == ''):
            print ('空路径') 
            pass
        elif ('.xls' in getpath):
            print('check',getpath)
            #self.path = open.getExistingDirectory()
            self.path = getpath
            self.pathLineEdit.setText(self.path)
        else:
            print ('所选文件不符合') 
            QMessageBox.information(self,                         #使用infomation信息框    
            "所选文件不符合",    
            "请选择.xls或.xlsx文件") 


    def test(self):
        print ('test')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    dialog = SelectDialog()
    if dialog.exec_():
        print ('dialog.exec_')
        pass
