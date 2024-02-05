


import os
import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5 import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import*
import pandas as pd
#from PyQt5 import QtCore, QtWidgets
import shutil
from datetime import datetime
import time
from make2 import *
import traceback
form_class = uic.loadUiType(f'./CLM_UI.ui')[0]
FROM_CLASS_Loading = uic.loadUiType("load.ui")[0]
user_name = os.getlogin()
cache_path = f'./cache/cache_{user_name}.csv'
if not os.path.isdir(os.path.dirname(cache_path)):
    os.makedirs(os.path.dirname(cache_path))

result_path = f'./result/{user_name}'
if not os.path.isdir(os.path.abspath(result_path)):
    os.makedirs(os.path.abspath(result_path))

isChanging = False
    
class WindowClass(QMainWindow, form_class) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)

        self.setWindowTitle("CheckListMaker 0.1")
        self.statusLabel = QLabel(self.statusbar)

        #self.setGeometry(1470,28,400,600)
        #self.setFixedSize(450,550)
        

        '''기본값입력■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■'''
        self.input_20.setText(os.path.abspath(result_path))

        self.import_cache_all()

        self.addcombo_xlsx_sheetnames(self.input_00.text(),self.combo_40)
        self.addcombo_xlsx_sheetnames(self.input_10.text(),self.combo_50)
        self.addcombo_xlsx_colnames(self.input_00.text(),self.combo_40.currentText(),self.combo_51)
        # self.make_ref_info_dict()
        # #소스가 있는 폴더(r2m 쉐어 포인트)        
        # self.btn_sourcePath_1.clicked.connect(lambda : self.파일열기(self.input_sourcePath.text()))
        # self.btn_sheetName.clicked.connect(self.load_sheetnames)
        # self.combo_sheetName.currentTextChanged.connect(self.load_enablecolnames)
        # self.btn_execute.clicked.connect(self.execute)

        # self.load_sheetnames()
        '''버튼 상호작용 입력■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■'''
        self.btn_00.clicked.connect(lambda:self.setFilePath(self.input_00))
        self.btn_01.clicked.connect(lambda:self.파일열기(self.input_00.text()))
        self.btn_10.clicked.connect(lambda:self.setFilePath(self.input_10))
        self.btn_11.clicked.connect(lambda:self.파일열기(self.input_10.text()))

        self.btn_40.clicked.connect(lambda:self.addcombo_xlsx_sheetnames(self.input_00.text(),self.combo_40))
        self.btn_50.clicked.connect(lambda:self.addcombo_xlsx_sheetnames(self.input_10.text(),self.combo_50))
        
        self.combo_40.currentTextChanged.connect(lambda:self.addcombo_xlsx_colnames(self.input_00.text(),self.combo_40.currentText(),self.combo_51))
        self.combo_40.currentTextChanged.connect(lambda:self.display_excel_data(self.input_00.text(),self.combo_40.currentText(),self.preview_table_2))
        
        self.combo_50.currentTextChanged.connect(lambda:self.display_excel_data(self.input_10.text(),self.combo_50.currentText(),self.preview_table))

        self.btn_20.clicked.connect(lambda:self.setDirectoryPath(self.input_20))
        self.btn_21.clicked.connect(lambda:self.파일열기(self.input_20.text()))
    

        self.preview_table.currentItemChanged.connect(lambda:print("1"))
        self.preview_table.currentCellChanged.connect(lambda:print("2"))
        self.preview_table.cellChanged.connect(lambda row, column: self.save_to_excel(row, column))
        #self.preview_table.currentItemChanged.connect(lambda:print("4"))

        self.btn_execute.clicked.connect(self.execute)
    def addcombo_xlsx_sheetnames(self,xlsx_filename,combo_object):
        #xlsx_filename = self.input_sourcePath.text()

       # global xls
        try:
            #xls = pd.read_excel(xlsx_filename, sheet_name=None)
            xls = pd.ExcelFile(xlsx_filename)
        except :
            self.popUp(f'엑셀파일을 먼저 선택해야 합니다.')
            return
        
        #sheet_names = xls.keys() if isinstance(xls, dict) else xls.sheet_names
        sheet_names = xls.sheet_names
        combo_object.clear()
        for sheet_name in sheet_names:
            combo_object.addItem(sheet_name)
    def addcombo_xlsx_colnames(self,xlsx_filename,xlsx_sheetname,combo_object):
        try:
            xls = pd.read_excel(xlsx_filename, sheet_name=xlsx_sheetname)
            #xls = pd.ExcelFile(xlsx_filename,sheet_name == xlsx_sheetname)
            
        except :
            self.popUp(f'엑셀파일을 먼저 선택해야 합니다.')
            return
        
        #sheet_names = xls.keys() if isinstance(xls, dict) else xls.sheet_names
        #sheet_names = xls.sheet_names
        names = xls.columns.tolist()
        combo_object.clear()
        for name in names:
            combo_object.addItem(name)


        #self.load_enablecolnames()
        #self.apply_colname(self.combo_sheetName.currentText())
            
                
    def display_excel_data(self,excel_file, sheet_name, qtablewidget):
        
        # Read Excel data
        try:
            df = pd.read_excel(excel_file, sheet_name)
        except Exception as e:
            print(f"Error reading Excel data: {e}")
            return
        

        global isChanging
        isChanging = True
        # Clear existing data in QTableWidget
        qtablewidget.clear()

        # Set column headers
        qtablewidget.setColumnCount(len(df.columns))
        qtablewidget.setHorizontalHeaderLabels(df.columns)

        # Set row count
        qtablewidget.setRowCount(len(df))

        # Populate QTableWidget with data
        for row in range(len(df)):
            for col in range(len(df.columns)):
                item = QTableWidgetItem(str(df.iloc[row, col]))
                            # Set font size to 10 (you can adjust the size as needed)
                font = QFont()
                font.setPointSize(8)
                item.setFont(font)
                qtablewidget.setItem(row, col, item)

        # Resize columns to content
        qtablewidget.resizeColumnsToContents()
        qtablewidget.resizeRowsToContents()

        isChanging = False

    def load_enablecolnames(self):
        
        xlsx_filename = self.input_sourcePath.text()
        sheet_name = self.combo_sheetName.currentText()

        # xls의 sheet_name의 열 이름을 불러와서 입력하는 코드 추가
        try:    
            df = xls[sheet_name] if isinstance(xls, dict) else pd.read_excel(xlsx_filename, sheet_name=sheet_name)
        except: 
            return
            
        # Filter out "Unnamed" columns
        valid_col_names = [col for col in df.columns if not col.startswith('Unnamed')]
    
        col_names = ', '.join(valid_col_names)
        
        self.input_enableColName.clear()
        self.input_enableColName.insertPlainText(col_names)

        self.apply_colname(sheet_name)

    def make_ref_info_dict(self):
        global ref_info_dict
        ref_info_dict = {}
        with open("ref_info.txt", "r", encoding='utf-8') as file:
            for line in file:
                parts = line.strip().split(',')
                sheet_name = parts[0]
                col_names = [col.strip() for col in parts[1:]]
                ref_info_dict[sheet_name] = col_names

    def apply_colname(self,cur_sheet_name):

        #cur_sheet_name이 ref_info_dic의 키와 일치하는게 있으면, 아래의 코드를 실행함
        if cur_sheet_name in ref_info_dict:
            print(cur_sheet_name)
            col_names = ref_info_dict[cur_sheet_name]

            self.input_mainColName.setText(col_names[0])
            self.input_targetColName.setText(','.join(col_names[1:]))
        # else:
        #     # Handle the case where cur_sheet_name is not found in ref_info_dic
        #     pass
        #일치하는게 없으면 아무것도안함.

    def execute(self):
        try:
            data_file = self.input_00.text()
            template_file = self.input_10.text()#self.combo_sheetName.currentText()
            data_sheet_name = self.combo_40.currentText()#os.path.join(self.input_resultPath.text(),f"{sheet_name}_{time.strftime('%y%m%d_%H%M%S')}.xlsx")
            sheet_name = self.combo_50.currentText()#self.input_mainColName.text()
            key_column = self.combo_51.currentText()#self.input_targetColName.text().split(',')
            result_path = self.input_20.text()
            cur_time = time.strftime('%Y%m%d_%H%M%S')
            result_file_name = os.path.join(result_path, f"{sheet_name}_{cur_time}.xlsx")
            

            self.loading = loading(self)
            #process_data_template(data_file, template_file, data_sheet_name, sheet_name, key_column)
            #self.worker_thread = WorkerThread(myWindow,sheet_name,input_file, output_file, criterion, required_parts)
            self.worker_thread = WorkerThread(myWindow,data_file,template_file, data_sheet_name, sheet_name, key_column, result_file_name)
            self.worker_thread.finished.connect(self.cleanup)
            self.worker_thread.start()
            
            #create_checklist(sheet_name,input_file, output_file, criterion, required_parts)

        except Exception as e:
            
            self.popUp(desText= traceback.format_exc())
            print(f'생성실패 : {e}')

    def cleanup(self):
        self.worker_thread = None

    def start_loading(self,qma):
        loading(self)

    def make_process(self,a,b,c,d,e,f):
        process_data_template(a,b,c,d,e,f)

        self.worker_thread.finished.emit()
        self.worker_thread.quit()
        self.loading
        self.loading.deleteLater()

        if self.check_0.isChecked():
            os.startfile(f)

    def 파일열기(self,filePath):
        try:
            os.startfile(os.path.abspath(filePath))
        except : 
            self.popUp(desText="파일 없음 : "+filePath)    

    def get_latest_file_in_directory(self, source_path, target_file):

        def find_latest_file(folder):
            latest_file = None
            latest_time = datetime.min

            for root, dirs, files in os.walk(folder):
                if target_file in files:
                    file_path = os.path.join(root, target_file)
                    modified_time = datetime.fromtimestamp(os.path.getmtime(file_path))
                    if modified_time > latest_time:
                        latest_file = file_path
                        latest_time = modified_time

            return latest_file

        latest_file_path = find_latest_file(source_path)
        return latest_file_path
        # if latest_file_path:
        #     # 파일 실행 코드 작성
        #     print(f"가장 최신의 파일 실행: {latest_file_path}")
        #     os.startfile(os.path.normpath(latest_file_path))

        # else:
        #     print(f"'{target_file}' 파일을 찾을 수 없습니다.")


    def find_folders_by_name(self, source_path, folder_name):
        #matching_folders = []
        
        for root, dirs, files in os.walk(source_path):
            if folder_name in dirs:
                folder_path = os.path.join(root, folder_name)
                #matching_folders.append(folder_path)
                return folder_path
            
        print(f"'{folder_name}' 이름을 가진 폴더를 찾을 수 없습니다.")
        

    source_path = fr'C:\Users\mssung\OneDrive - Webzen Inc\R2M_Build\KR'
    #folder_name = 'YourFolderName'  # 검색할 폴더 이름

    # matching_folders = find_folders_by_name(source_path, folder_name)

    # if matching_folders:
    #     for folder in matching_folders:
    #         print(f"폴더 경로: {folder}")
    # else:
    #     print(f"'{folder_name}' 이름을 가진 폴더를 찾을 수 없습니다.")


    def print_log(self, log): # / - \ / - \ / ㅡ ㄷ
        self.progressLabel.setText(log)
        QApplication.processEvents()
        
    def popUp(self,desText,titleText="error"):
        msg = QMessageBox()  
        #msg.setGeometry(1520,28,400,2000)
        msg.setText(desText)
        msg.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)

        x = msg.exec_()

    def setFilePath(self,target):
        path = QFileDialog.getOpenFileName(self)
        if path != "" :
            target.setText(path[0])
            #dirname = os.path.dirname(path[0])
            #cur_dirname = os.getcwd()
            #print(dirname,cur_dirname)
            # if f'{dirname}' == f'{cur_dirname}' :
            #     target.setText(os.path.basename(path[0])) 
            # else:
            #     target.setText(path[0])
    def setDirectoryPath(self, target):
        path = QFileDialog.getExistingDirectory(self)
        if path != "" :
            target.setText(path)

    def save_to_excel(self, row, column):
        if isChanging :
            return
        file_name = self.input_10.text()
        sheet_name = self.combo_50.currentText()

        # Get the new value from the QTableWidget
        new_value = self.preview_table.item(row, column).text()

        # Load existing Excel file with openpyxl
        book = load_workbook(file_name)

        # Access the active sheet
        sheet = book[sheet_name]

        # Convert 0-based QTableWidget coordinates to 1-based Excel coordinates
        excel_row = row + 2
        excel_column = column + 1

        # Update the value in the Excel sheet
        try:
            sheet.cell(row=excel_row, column=excel_column, value=new_value)
        except:
            return

        # Save changes back to the Excel file
        book.save(file_name)

        print("save_to_excel")

    def import_cache_all(self, any_widget = None):
        '''any_widget : [QLineEdit,'input_00']'''
        try:
            # Load CSV file with tab delimiter and utf-16 encoding
            df = pd.read_csv(cache_path, sep='\t', encoding='utf-16', index_col='key')
            if any_widget == None :
                all_widgets = self.findChildren((QLineEdit,  QComboBox, QCheckBox, QPlainTextEdit, QDateEdit))
            else:
                all_widgets = [self.findChild(any_widget[0] ,any_widget[1])]

            for widget in all_widgets:
                object_name = widget.objectName()
                if object_name in df.index:
                    value = str(df.loc[object_name, 'value'])
                    if isinstance(widget, (QLineEdit,QLabel,QPushButton)):
                        widget.setText(value)
                    elif isinstance(widget, QComboBox):
                        # Set selected index based on the value, adjust as needed
                        index = widget.findText(value)
                        if index != -1:
                            widget.setCurrentIndex(index)
                    elif isinstance(widget, QCheckBox):
                        widget.setChecked(value.lower() == 'true')
                    elif isinstance(widget, QPlainTextEdit):
                        widget.setPlainText(value)
                    elif isinstance(widget, QDateEdit):
                        date_format = "yyyy-MM-dd"
                        date = QDate.fromString(value, date_format)
                        widget.setDate(date)

        except Exception as e:
            print(f"Error importing cache: {e}")

    def export_cache_all(self):
        try:
            data = {'key': [], 'value': []}

            all_widgets = self.findChildren((QLineEdit,  QComboBox, QCheckBox, QPlainTextEdit,QDateEdit))

            for widget in all_widgets:
                value = ""
                if isinstance(widget, (QLineEdit,QLabel,QPushButton,QDateEdit)) :
                    value = widget.text()
                elif isinstance(widget, QComboBox):
                    value = widget.currentText()
                elif isinstance(widget, QCheckBox):
                    value = str(widget.isChecked())
                elif isinstance(widget, QPlainTextEdit):
                    value = widget.toPlainText()

                if value != "":
                    key = widget.objectName()
                    data['key'].append(key)
                    data['value'].append(value)

            df = pd.DataFrame(data)
            df.set_index('key', inplace=True)
            df.to_csv(cache_path, sep='\t', encoding='utf-16')
        except Exception as e:
            print(f"Error exporting cache: {e}")

    def closeEvent(self,event):
        print("end")

        #self.export_cache(isForced= True)
        self.export_cache_all()

class loading(QWidget,FROM_CLASS_Loading):
    
    def __init__(self,parent):
        super(loading, self).__init__(parent)    
        self.setupUi(self) 
        #self.resize(parent.size())
        self.setFixedSize(parent.size())
        self.center()
        # Get the size of the parent widget and set the loading widget to the same size
        
        self.show()
        
        self.movie = QMovie('lcu_ui_ready_check.gif', QByteArray(), self)
        self.movie.setCacheMode(QMovie.CacheAll)
        self.label.setMovie(self.movie)
        self.label.setScaledContents(True)
        #self.movie.set(500,500)
        self.movie.start()
        self.setWindowFlags(Qt.FramelessWindowHint)
    # 위젯 정중앙 위치
    def center(self):
        
        size=self.size()
        ph = self.parent().geometry().height()
        pw = self.parent().geometry().width()
        self.move(int(pw/2 - size.width()/2), int(ph/2 - size.height()/2))
        self.move(int(pw/2 - size.width()/2), int(ph/2 - size.height()/2))
        
class WorkerThread(QThread):
    finished = pyqtSignal()

    def __init__(self,window, a,b, c, d, e,f):
        super().__init__()
        
        self.window = window
        self.a = a
        self.b = b
        self.c = c
        self.d = d
        self.e = e
        self.f = f

    def run(self):
        #create_checklist(sheet_name,input_file, output_file, criterion, required_parts)
        #create_checklist(self.a,self.b, self.c, self.d, self.e)
        self.window.make_process(self.a, self.b, self.c, self.d, self.e, self.f)
        self.finished.emit()

if __name__ == "__main__" :
    #QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv) 

    #WindowClass의 인스턴스 생성
    myWindow = WindowClass() 

    #프로그램 화면을 보여주는 코드
    myWindow.show()

    #프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec_()