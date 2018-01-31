from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import sys,os,json
from punchcard import checkExcel ,creatExcel

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
        
        




class SelectDialog(QDialog):
    def __init__(self, parent=None):
        super(SelectDialog, self).__init__(parent)
        self.info = PersonalInfo.open()
        self.labrec = ''
        self.path = os.getcwd()[0]
        self.initUI()
        self.setWindowTitle("批补卡")
        self.resize(340, 100)
        # print self.info

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
        self.pathLineEdit.setFixedWidth(400)
        self.pathLineEdit.setText(self.path)

        self.pathLineEdit.setDragEnabled = True
        self.pathLineEdit.setDropEnabled = True
        # self.pathLineEdit.dragLeaveEvent()

        # self.pathLineEdit.connect(self.pathLineEdit, QtCore.SIGNAL("dropped"), self.test)



        grid.addWidget(self.pathLineEdit, 0, 1)
        button = QPushButton("更改")
        button.clicked.connect(self.changePath)
        grid.addWidget(button, 0, 2)


        # grid.addWidget(QLabel("导出："), 1, 0)
        # self.pathLineEdit2 = QLineEdit()
        # self.pathLineEdit2.setFixedWidth(300)
        # self.pathLineEdit2.setText(self.path)
        # grid.addWidget(self.pathLineEdit2, 1, 1)
        # button2 = QPushButton("更改")
        # button2.clicked.connect(self.changePath)
        # grid.addWidget(button2, 1, 2)

        self.tfnum = QLineEdit()
        self.tfnum.setFixedWidth(100)
        self.tfnum.setText(worknum)
        self.tfnum.move(100,50)
        grid.addWidget(QLabel("工号："), 2, 0)
        grid.addWidget(self.tfnum, 2, 1)

        button3 = QPushButton("查看")
        button3.clicked.connect(self.check)
        grid.addWidget(button3, 2, 2)

        
        # buttonBox = QDialogButtonBox()
        # buttonBox.setOrientation(Qt.Horizontal)  # 设置为水平方向
        # buttonBox.setStandardButtons(QDialogButtonBox.Ok|QDialogButtonBox.Cancel)
        # buttonBox.accepted.connect(self.accept)  # 确定
        # buttonBox.rejected.connect(self.reject)  # 取消
        # grid.addWidget(buttonBox, 4, 1)
        self.setLayout(grid)

    def check(self):
        result = checkExcel(self.path,self.tfnum.text())
        print ('路径:',self.path,'工号：',self.tfnum.text(),'条数：',len(result[1]))
        # print (result)

        self.listw = QListWidget()
        self.listw.addItems(result[0])
        self.listw.setFixedWidth(400)
        self.grd.addWidget(self.listw, 3, 1)

        # print ()
        if self.labrec == '' :
            self.labrec = QLabel()
            self.labrec.setText("有%d条未打卡记录" % len(result[1]))
            self.grd.addWidget(self.labrec, 4, 1)
        else:
            self.labrec.setText("有%d条未打卡记录" % len(result[1]))

        # self.grd.addWidget(QLabel("具体地点："), 5, 0)
        # self.grd.addWidget(QLineEdit(""), 5, 1)
        
        self.grd.addWidget(QLabel("工作原因："), 6, 0)
        self.grd.addWidget(QLineEdit(""), 6, 1)

        button = QPushButton("查看")
        button.clicked.connect(self.output)
        self.grd.addWidget(button, 6, 2)
        
        if len(result)>2:
            self.grd.addWidget(QLabel("个人原因："), 7, 0)
            self.grd.addWidget(QLineEdit(""), 7, 1)


    def output(self):
        
        open = QFileDialog()
        self.outputpath = open.getExistingDirectory()
        self.grd.addWidget(QLabel("导出路径："), 8, 0)
        self.grd.addWidget(QLineEdit(self.outputpath), 8, 1)

    def test(self):
        print ('test')
        # self.grd.addWidget(QLabel("工作原因："), 3, 0)
        # self.grd.addWidget(QLineEdit("test"), 3, 1)

        # self.grd.addWidget(QLabel("个人原因："), 3, 2)
        # self.grd.addWidget(QLineEdit("test"), 3, 3)

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


# from PyQt5 import QtWidgets
# from PyQt5.QtWidgets import QFileDialog

# class MyWindow(QtWidgets.QWidget):
#     def __init__(self):
#         super(MyWindow,self).__init__()
#         self.myButton = QtWidgets.QPushButton(self)
#         self.myButton.setObjectName("myButton")
#         self.myButton.setText("Test")
#         self.myButton.clicked.connect(self.msg)

#     def msg(self):
#         directory1 = QFileDialog.getExistingDirectory(self,
#                                     "选取文件夹",
#                                     "C:/")                                 #起始路径
#         print(directory1)

#         fileName1, filetype = QFileDialog.getOpenFileName(self,
#                                     "选取文件",
#                                     "C:/",
#                                     "All Files (*);;Text Files (*.txt)")   #设置文件扩展名过滤,注意用双分号间隔
#         print(fileName1,filetype)

#         files, ok1 = QFileDialog.getOpenFileNames(self,
#                                     "多文件选择",
#                                     "C:/",
#                                     "All Files (*);;Text Files (*.txt)")
#         print(files,ok1)

#         fileName2, ok2 = QFileDialog.getSaveFileName(self,
#                                     "文件保存",
#                                     "C:/",
#                                     "All Files (*);;Text Files (*.txt)")

# if __name__=="__main__":  
#     import sys  
  
#     app=QtWidgets.QApplication(sys.argv)  
#     myshow=MyWindow()
#     myshow.show()
#     sys.exit(app.exec_())  

