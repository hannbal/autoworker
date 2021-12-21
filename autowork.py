# -*- coding: utf-8 -*-
"""
Created on Mon Dec 13 16:27:35 2021

@author: tony
"""

#自动执行任务
import pandas as pd
import pyautogui as pag
import pyperclip
import  os.path, sys,glob
from PyQt5.QtWidgets import QApplication,QMainWindow,QWidget,QTableWidget,QGridLayout,QHBoxLayout,QVBoxLayout,QTableWidgetItem
from PyQt5.QtWidgets import QLabel,QGroupBox,QListWidget,QPushButton,QTabWidget,QListWidgetItem,QHeaderView,QComboBox,QSpinBox,QRadioButton

from PyQt5.QtGui import QStandardItemModel, QStandardItem
from PyQt5.QtCore import Qt,QModelIndex 
from pynput import keyboard
import time
import _thread

class mytab(QTabWidget):
    def __init__(self, *args, **kwargs):
        super(mytab,self).__init__(*args, **kwargs)
    
    
class myqt(QMainWindow):
    def __init__(self, *args, **kwargs):
        self.currectPath =os.getcwd()
        self.file_list={}  
        self.desktop = QApplication.desktop()
        self.screenRect = self.desktop.screenGeometry()
        self.op_list=["单击","双击","右击","移动、悬停","输入","回车","滚轮","热键","windows命令行","等待"]
        # self.pd_data={}
        
        super(myqt,self).__init__(*args, **kwargs)
        self.w_height = self.screenRect.height()
        self.w_width = self.screenRect.width()
#主体窗口        
        self.setMinimumSize(1024, 800)
        self.setWindowTitle("autowork 1.0")
    #嵌套第一层窗口           
        self.centralwidget = QWidget()
        self.setCentralWidget(self.centralwidget)        
        self.centralwidget.setStyleSheet("QWidget{background-color: black;font: 16px;color: white;}")     
    #调用窗口初始化   
        self.setupUi()
        
    def scandirs(self):
        for filename in glob.glob(self.currectPath+'\\*.csv'):   
            filename=filename.replace('//','/').replace('/', '\\')                 
            self.file_list[filename.split('\\')[-1][:-4]]=filename
        

    

                    
    def setupUi(self):
        self.layout= QGridLayout()
        #边界
        self.layout.setContentsMargins(20,20,20,20)
        #网格间距
        self.layout.setSpacing(10)
        
# #窗口top
        self.layout_top= QHBoxLayout()
        self.label_name = QLabel("使用说明",self)
        self.layout_top.addWidget(self.label_name)
        self.label_text = QLabel("单击,双击,右击,移动、悬停,输入,回车,滚轮,热键,windows命令行,等待,运行前请先保存！",self)
        self.label_text.setWordWrap(True)
        self.label_text.setStyleSheet("QLabel{padding:8px}")
        self.layout_top.addWidget(self.label_text)
        
        self.layout_top_right = QVBoxLayout()
        self.box_top_right = QGroupBox()
        self.box_top_right.setLayout(self.layout_top_right)
        self.layout_top.addWidget(self.box_top_right)
        
#运行指定次数控件    
        self.radio_run = QRadioButton("指定次数运行")
        self.radio_run.setObjectName("radio_run")
        self.radio_run.setChecked(True)
        self.layout_top_right.addWidget(self.radio_run) 
        
        self.edit_text = QSpinBox()
        self.edit_text.setMinimum(1)
        self.layout_top_right.addWidget(self.edit_text)
        self.pb_run = QPushButton("运行")
        self.pb_run.setStyleSheet("QPushButton{background-color: #588bb1;font: 16px;color: blue;border-color:white;border-width:10px;}")
        self.layout_top_right.addWidget(self.pb_run)


#循环运行控件
        self.radio_runalltime = QRadioButton("循环运行")
        self.radio_runalltime.setObjectName("radio_runalltime")
        self.layout_top_right.addWidget(self.radio_runalltime) 
        
        self.pb_runalltime = QPushButton("循环运行")
        self.pb_runalltime.setStyleSheet("QPushButton{background-color: #588bb1;font: 16px;color: blue;border-color:white;border-width:10px;}")
        self.pb_runalltime.setEnabled(False)
        self.layout_top_right.addWidget(self.pb_runalltime)        
        
    #两个标签 列宽比例设置    
        self.layout_top.setStretch(0,1)
        self.layout_top.setStretch(1,10)
        self.layout_top.setStretch(2,1)
    #添加到主窗口
        self.tab_top = mytab()
        self.tab_top.setLayout(self.layout_top)
        self.layout.addWidget(self.tab_top)
        

# #窗口middle       
    #左边文件列表
        self.layout_middle = QHBoxLayout()
        self.file_list_box = QGroupBox()
        self.file_list_box.setTitle('文件列表')
        
        self.file_list_view = QListWidget(self.file_list_box)
        self.layout_list =QVBoxLayout()
        self.file_list_view.setStyleSheet("QListWidget::item:selected{background-color:rgb(155,194,230)}")
       
        #按钮        
        self.pb_newfile = QPushButton("新增文件")
        self.pb_newfile.setStyleSheet("QPushButton{background-color: #588bb1;font: 16px;color: blue;border-color:white;border-width:10px;}")
        self.pb_savefile = QPushButton("保存当前文件")
        self.pb_savefile.setStyleSheet("QPushButton{background-color: #588bb1;font: 16px;color: blue;border-color:white;border-width:10px;}")
        self.pb_newop = QPushButton("新增操作")
        self.pb_newop.setStyleSheet("QPushButton{background-color: #588bb1;font: 16px;color: blue;border-color:white;border-width:10px;}")
        self.pb_delop = QPushButton("删除操作")
        self.pb_delop.setStyleSheet("QPushButton{background-color: #588bb1;font: 16px;color: blue;border-color:white;border-width:10px;}")
        
       
        self.layout_list.addWidget(self.file_list_view)

        self.layout_list.addWidget(self.pb_newop)
        self.layout_list.addWidget(self.pb_delop)
        self.layout_list.addWidget(self.pb_newfile)
        self.layout_list.addWidget(self.pb_savefile)
        self.file_list_box.setLayout(self.layout_list)
        
        
    #右边操作按记录   
        self.op_list_box = QGroupBox()
        self.op_list_box.setTitle('操作列表')
        
        self.op_table_view = QTableWidget()
        self.op_table_header = ["操作","命令","备注"]
        self.op_table_view.setColumnCount(3)
        self.op_table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.op_table_view.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.op_table_view.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)

        self.op_table_view.horizontalHeader().setStyleSheet("QHeaderView::section{background-color:rgb(155,194,230)}")
        

        self.op_table_view.setHorizontalHeaderLabels(self.op_table_header)        
        self.layout_table = QVBoxLayout()
    
        self.layout_table.addWidget(self.op_table_view)
        self.op_list_box.setLayout(self.layout_table)
        
    #下方左右两个加入下方总layout
        self.layout_middle.addWidget(self.file_list_box)
        self.layout_middle.addWidget(self.op_list_box)
        
    #middle两个部分 列宽比例设置     
        self.layout_middle.setStretch(0,1)
        self.layout_middle.setStretch(1,3)
    #添加到主窗口
        self.tab_middle = mytab()
        self.tab_middle.setLayout(self.layout_middle)
        self.layout.addWidget(self.tab_middle)        


# #窗口bottom       
        # self.tab_bottom = mytab(self)
        # self.layout.addWidget(self.tab_bottom)
        
        
        self.centralwidget.setLayout(self.layout)
#top middle bottom 行宽比例设置
        self.layout.setRowStretch(0,3)
        self.layout.setRowStretch(1,10)

#按钮关联        
        self.pb_savefile.clicked.connect(self.save_file)
        self.pb_newfile.clicked.connect(self.new_file)   
        self.file_list_view.clicked.connect(self.listview_changeevent)
        self.pb_newop.clicked.connect(self.table_add)
        self.pb_delop.clicked.connect(self.table_del)
        self.pb_run.clicked.connect(self.run)
        
        self.radio_run.clicked.connect(self.radio_change)
        self.radio_runalltime.clicked.connect(self.radio_change)
        self.pb_runalltime.clicked.connect(self.run_alltime)
        
#全屏启动             
        self.isFullScreen()
        self.showMaximized()  
        self.scandirs()
        self.file_list_view.clear()
        for file in self.file_list.keys():
            item = QListWidgetItem()
            item.setText(file)
            self.file_list_view.addItem(item) 
        
        

    def pd_readcsv(self,path):
        try:
            df = pd.read_csv(path,encoding="utf-8")
            if df.size==0:
                df = pd.DataFrame([['1','','']])
        except Exception:
            df = pd.DataFrame([[' ',' ',' '],[' ',' ',' '],[' ',' ',' ']])
        self.pd_data = df
        if any(df):
            
            start_row = 0
            df_columns = df.shape[1]
            df_rows = df.shape[0]
            df_header = df.columns.values.tolist()
            self.op_table_view.setColumnCount(df_columns)
            self.op_table_view.setHorizontalHeaderLabels(df_header)
            
            for row in range(0,df_rows):
                self.op_table_view.setRowCount(row+1)
                for column in range(df_columns):
                    value= ''
                    if row < df_rows:
                        value ='' if None==(df.iloc[row,column]) else str(df.iloc[row,column])
                    if column ==0:
                        tempItem = QComboBox()
                        tempItem.addItems(self.op_list)
                        
                        if value.isnumeric():
                            tempItem.setCurrentIndex(int(value))
                        self.op_table_view.setCellWidget((row - start_row),column,tempItem)
                    else:
                        tempItem = QTableWidgetItem()
                        self.op_table_view.setItem((row - start_row),column,tempItem)
                        self.op_table_view.item((row - start_row),column).setText(value)#.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter())
        
  
    def listview_changeevent(self):
        index = self.file_list_view.currentItem().text()
        self.pd_readcsv(self.file_list[index])
        print("切换至'{0}'".format(self.file_list_view.currentItem().text()))
        
    
    def new_file(self,event):
        # sender = self.sender()
        filepath = self.currectPath+'\\'+'temp'+str(time.localtime().tm_mon)+'-'+str(time.localtime().tm_mday)+'-'+str(time.localtime().tm_hour)+str(time.localtime().tm_min)+str(time.localtime().tm_sec)+".csv"
        
        self.file_list[filepath.split('\\')[-1]]=filepath
        print('make a file :{0}'.format(filepath))
        file = open(filepath,'w',encoding="utf-8",newline='\r')
        file.writelines("操作,命令,备注\r")
        file.writelines("1,,")
        file.close()

        item = QListWidgetItem()
        item.setText(filepath.split('\\')[-1])
        self.file_list_view.addItem(item) 
    
    def save_file(self):
        if None==self.file_list_view.currentItem():
            print('没选择文件')
            return 0
        filename = self.file_list_view.currentItem().text()
        filepath = self.file_list[filename]
        list=[]
        for x in range(self.op_table_view.rowCount()):
            list.append([self.op_table_view.cellWidget(x,0).currentIndex(),self.op_table_view.item(x,1).text(),self.op_table_view.item(x,2).text()])
        temppd=pd.DataFrame(list,columns=["操作","命令","备注"])
        
        
        try:
            temppd.to_csv(filepath,index=0)
            print('保存成功')
        except Exception:
            print('file-saving failed')
            
            
    def table_add(self):
        new_row =self.op_table_view.rowCount()
        self.op_table_view.setRowCount(new_row+1)
        
        tempItem = QComboBox()
        tempItem.addItems(self.op_list)
        self.op_table_view.setCellWidget(new_row,0,tempItem)
        
        for x in range(1,3):    
            tempItem = QTableWidgetItem()
            self.op_table_view.setItem(new_row,x,tempItem)
            self.op_table_view.item(new_row,x).setText(' ')
        
        

    def table_del(self):
        
        row = self.op_table_view.currentRow()
        if row<0:
            print('没有操作可以删除')
        else:    
            self.op_table_view.removeRow(row)
    
    def run(self):
        def click(pngfile):
            #单击
            a=pag.locateOnScreen(pngfile,confidence=0.9)
            if a is None:
                raise Exception('没找到图片,已停止运行！')
            try:
                pag.click(a)
                print('鼠标单击 : {0}'.format(pngfile))
            except Exception as e:
                print(e)

        def doubleclick(pngfile):
            #双击
            a=pag.locateOnScreen(pngfile,confidence=0.9)
            if a is None:
                raise Exception('没找到图片,已停止运行！')
            try:
                pag.doubleClick(a)
                print('鼠标双击 : {0}'.format(pngfile))
            except Exception as e:
                print(e)


        def move(pngfile):
            #移动、悬停
            a=pag.locateOnScreen(pngfile,confidence=0.9)
            if a is None:
                raise Exception('没找到图片,已停止运行！')
            try:
                pag.moveTo(a)
                print('鼠标移动到 : {0}'.format(pngfile))
            except Exception as e:
                print(e)
        
        def rightclick(pngfile):
            #右击
            a=pag.locateOnScreen(pngfile,confidence=0.9)
            if a is None:
                raise Exception('没找到图片,已停止运行！')
            try:
                pag.rightClick(a)
                print('鼠标移动到 : {0}'.format(pngfile))
            except Exception as e:
                print(e)

        def input_text(text):
            #输入
            pyperclip.copy(text)
            pag.hotkey('ctrl','v')
            print('输入文本:{0}'.format(text))
 
        def scoll():
            #滚轮
            print('scoll')

        def hotkey(keylist):
            #热键
            try:
                for key in keylist:
                    pag.keyDown(key)                
                for key in keylist:
                    pag.keyUp(key) 
                print('执行快捷键 : {0}'.format(keylist))
            except Exception as e:
                print(e)

        def windows_cmd():
            #window
            print('windows_cmd')


        def wait_for(sec):
            #等待
            time.sleep(sec)
            print('wait_for {0} secend'.format(sec))
    
        rowcount =self.op_table_view.rowCount()
        if rowcount==0:
            print("请先添加操作")
            return 0
        self.save_file() #运行自动保存
        for times in range(self.edit_text.value()):
            for row in range(rowcount):
                actioncode = self.op_table_view.item(row,1).text()
                if self.op_table_view.cellWidget(row,0).currentIndex() ==0:
                    try:
                        click(actioncode.strip())
                    except Exception as e:
                        print(e)
                        break
              
                elif self.op_table_view.cellWidget(row,0).currentIndex() ==1:
                    # pngfile=self.op_table_view.item(row,1).text()
                    try:
                        doubleclick(actioncode.strip())
                    except Exception as e:
                        print(e)
                        break
                
                elif self.op_table_view.cellWidget(row,0).currentIndex() ==2:
                    # pngfile=self.op_table_view.item(row,1).text()
                    try:
                        rightclick(actioncode.strip())
                    except Exception as e:
                        print(e)
                        break    
                
                elif self.op_table_view.cellWidget(row,0).currentIndex() ==3:
                    # pngfile=self.op_table_view.item(row,1).text()
                    try:
                        move(actioncode.strip())
                    except Exception as e:
                        print(e)
                        break
                    
                elif self.op_table_view.cellWidget(row,0).currentIndex() ==4:
                    input_text(actioncode)
                
                elif self.op_table_view.cellWidget(row,0).currentIndex() ==5:
                    # pngfile=self.op_table_view.item(row,1).text()
                    try:
                        hotkey(['enter'])
                    except Exception as e:
                        print(e)
                        break
                
                elif self.op_table_view.cellWidget(row,0).currentIndex() ==6:
                    pag.scroll(-1000)
                    scoll()
                    
                elif self.op_table_view.cellWidget(row,0).currentIndex() ==7:
                    keylist = actioncode.split('-')
                    if keylist[0] == '':
                        print('请先输入快捷键，已停止程序')
                        break
                    try:
                        hotkey(keylist)
                    except Exception as e:
                        print(e)
                        break
                    
                    
                elif self.op_table_view.cellWidget(row,0).currentIndex() ==8:
                    windows_cmd()
                    
                elif self.op_table_view.cellWidget(row,0).currentIndex() ==9:
                    if self.op_table_view.item(row,1).text().isnumeric():
                        wait_for(int(self.op_table_view.item(row,1).text()))
                    else:
                        print('请输入正确的暂停秒数')
                else:
                    print('命令输入错误')
  
    def radio_change(self,event):
        sender = self.sender()
        if sender.objectName()=="radio_runalltime":
            self.edit_text.setEnabled(False)
            self.pb_run.setEnabled(False)
                 
            self.pb_runalltime.setEnabled(True)
        
        elif sender.objectName()=="radio_run":
            self.edit_text.setEnabled(True)
            self.pb_run.setEnabled(True)
                 
            self.pb_runalltime.setEnabled(False)
        else:
            pass
  
    def run_alltime(self):
        self.k_run =True
        def on_press(key):
            
            if key== keyboard.Key.esc :
                self.k_run =False
                return False
        def keyboardlistener(threadName, delay):
            while True:
                with keyboard.Listener(on_press=on_press) as listener:
                    listener.join()    
                break
        def keeprun(threadName, delay):

            while(self.k_run):
                print("开始循环执行")
                self.run()
            print("循环执行停止")
        
        try:
            _thread.start_new_thread( keyboardlistener, ("Thread-1", 2, ) )
            _thread.start_new_thread( keeprun, ("Thread-2", 3, ) )
        except Exception as e:
            print("Error: 无法启动线程，{0}".format(e))
            
           



        
class PandasTableModel(QStandardItemModel):
    def __init__(self, data, parent=None):
        QStandardItemModel.__init__(self, parent)
        self._data = data
        for row in data.values.tolist():
            data_row = [ QStandardItem("{}".format(x)) for x in row ]
            self.appendRow(data_row)
        return

    def rowCount(self, parent=None):
        return len(self._data.values)

    def columnCount(self, parent=None):
        return self._data.columns.size

    def headerData(self, x, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._data.columns[x]
        if orientation == Qt.Vertical and role == Qt.DisplayRole:
            return self._data.index[x]
        return None    
    def appendnewRow(self,data):
        position = self.rowCount()+1
        index = QModelIndex()
        self.beginInsertRows(index,position,position)
        
        self.endInsertColumns()
        self.dirty=True
        return True



        



    
if __name__=='__main__':
    app = QApplication(sys.argv)
    
    ui =myqt()
    sys.exit(app.exec_())
