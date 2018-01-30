from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import sys,os,json
from punchcard import checkExcel ,creatExcel


class PersonalInfo(object):
    """docstring for PersonalInfo"""
    def __init__(self, arg=None):
        super(PersonalInfo, self).__init__()
        # self.arg = arg
        # self.open()
    def open(test):
        model={} #存放读取的数据
        if os.access("./source/info.json", os.F_OK):
            with open("./source/info.json",'r',encoding='utf-8') as json_file:
                model=json.load(json_file)
            return model
        else:
            # return []
            pass
            # return '不存在文件'
        




class SelectDialog(QDialog):
    def __init__(self, parent=None):
        super(SelectDialog, self).__init__(parent)
        self.info = PersonalInfo().open()

        self.path = os.getcwd()
        self.initUI()
        self.setWindowTitle("补卡")
        self.resize(340, 100)
        # print self.info

    def initUI(self):
        worknum = ''
        if self.info:
            print (self.info)
            worknum = self.info['num']
        else: 
            print ('无info')
        grid = QGridLayout()
        self.grd = grid
        grid.addWidget(QLabel("考勤表："), 0, 0)
        self.pathLineEdit = QLineEdit()
        self.pathLineEdit.setFixedWidth(400)
        self.pathLineEdit.setText('')
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

        tfnum = QLineEdit()
        tfnum.setFixedWidth(100)
        tfnum.setText(worknum)
        tfnum.move(100,50)
        grid.addWidget(QLabel("工号："), 2, 0)
        grid.addWidget(tfnum, 2, 1)

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
        print ('路径:',self.path[0])
        result = checkExcel(self.path[0],'19489')
        # print (result)


        listw = QListWidget()
        listw.addItems(result[0])
        listw.setFixedWidth(400)
        self.grd.addWidget(listw, 3, 1)


    def test(self):
        self.grd.addWidget(QLabel("工作原因："), 3, 0)
        self.grd.addWidget(QLineEdit("test"), 3, 1)

        self.grd.addWidget(QLabel("个人原因："), 3, 2)
        self.grd.addWidget(QLineEdit("test"), 3, 3)

    def changePath(self):
        print ('enter')
        open = QFileDialog()
        self.path=open.getOpenFileName()
        print('check',self.path)
        #self.path = open.getExistingDirectory()
        self.pathLineEdit.setText(self.path[0])

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

