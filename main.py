from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import sys,os,json
from punchcard import checkExcel ,creatExcel
import sip


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
            # return []
            pass
            # return '不存在文件'
    
    def save(key,value):
        info = PersonalInfo.open()
        info[key] = value
        # print ('save info',info)
        with open("./source/info.json",'w',encoding='utf-8') as json_file:
            json.dump(info,json_file,ensure_ascii=False)

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
    def initUI(self):
        worknum = ''
        # kqpath = ''
        if self.info:
            if 'num' in self.info:
                worknum = self.info['num']
            if 'filepath' in self.info:
                self.path = self.info['filepath']

        else: 
            print ('无info')


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
        print ('路径:',self.path,'工号：',self.tfnum.text(),'条数：',len(result[1]))

        #未打卡记录表
        if self.record_detail is None :
            print ('no record_detail')
            self.record_detail = QListWidget()
            self.record_detail.addItems(result[0])
            # self.record_detail.setFixedWidth(400)
            # self.record_detail.setFixedHeight(150)
            self.grd.addWidget(self.record_detail, 3, 1)
        else:
            for i in range(self.record_detail.count())[::-1]:
                select_item = self.record_detail.takeItem(i)
                self.record_detail.removeItemWidget(select_item)
            self.record_detail.addItems(result[0])  


        # 未打卡数
        if self.record_num is None :
            self.record_num = QLabel()
            self.record_num.setText("有%d条未打卡记录" % len(result[1]))
            self.grd.addWidget(self.record_num, 4, 1)
        else:
            self.record_num.setText("有%d条未打卡记录" % len(result[1]))


        #工作原因
        if self.work_reason is None:
            self.work_reason = QLineEdit()
            self.grd.addWidget(QLabel("工作原因："), 6, 0)
            self.grd.addWidget(self.work_reason, 6, 1)



        if len(result[1])>2:
            # print ('记录大于2')
            if self.personal_reason is None:
                self.personal_reason = QLineEdit()
                self.personal_reason_tit = QLabel("个人原因：")
                self.grd.addWidget(self.personal_reason_tit, 7, 0)
                self.grd.addWidget(self.personal_reason, 7, 1)
        else:
            if self.personal_reason :
                print ('chekck')
                self.grd.removeWidget(self.personal_reason)
                self.grd.removeWidget(self.personal_reason_tit)

                sip.delete(self.personal_reason)
                sip.delete(self.personal_reason_tit)

                self.personal_reason = None
                self.personal_reason_tit = None

        if self.btn_output is None:
            self.btn_output = QPushButton("设置")
            self.btn_output.clicked.connect(self.output)
            self.grd.addWidget(QLabel("导出路径："), 8, 0)   
            self.grd.addWidget(self.btn_output, 8, 2)


    def output(self):
        open = QFileDialog()
        if self.outputpath is None: 
            print ('set output')    
            self.outputpath = QLabel()
            self.grd.addWidget(self.outputpath, 8, 1)
        outputpath_str = open.getExistingDirectory()
        self.outputpath.setText(outputpath_str)

    def test(self):
        print ('test')


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
            PersonalInfo.save('filepath',self.path)
        else:
            print ('所选文件不符合') 
            QMessageBox.information(self,                         #使用infomation信息框    
            "所选文件不符合",    
            "请选择.xls或.xlsx文件") 

if __name__ == '__main__':
    app = QApplication(sys.argv)
    dialog = SelectDialog()
    if dialog.exec_():
        print ('dialog.exec_')
        pass
