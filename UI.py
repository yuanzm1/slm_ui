from designer_ui import Ui_MainWindow
from datetime import datetime
import configparser, csv
import sys,os
import re
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QMainWindow, QDialog, QGraphicsScene, \
    QLabel, QMessageBox, QTableWidget, QHeaderView, QTableWidgetItem
from PyQt5.QtGui import QColor

def read_ini_file(filename, encoding='utf-8'):
    config = configparser.ConfigParser(comment_prefixes='/', allow_no_value=True)
    config.read(filename, encoding)
    return config

class Myui(QMainWindow,Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.tabWidget.setCurrentIndex(0)
        self.check_config_ini_path()
        self.config_ini = read_ini_file(os.path.join(os.getcwd(), 'Config.ini'), encoding='utf-8-sig')  # config_ini
        self.combox_index = 0
        self.model_ini_path = self.get_model_ini_path()
        self.config = read_ini_file(self.model_ini_path, encoding='utf-8-sig')  # model_ini
        self.labels = [self.findChild(QLabel, f"label_{i}") for i in range(98, 148)]
        self.innerQty = 20  # 内包装数量
        self.check_num = 0
        self.save_title_bool = False
        self.model_param, self.model_scan = '', ''
        self.total_check = 0  # 累计集合包装完成数量
        self.box_num = 0  # 内箱编号
        self.platform_num = 0  # 内箱台次
        self.save_path = ''
        self.init_flag_save_path = 0
        self.pre_sncode = []
        self.set_window_title()
        self.init_combobox()
        self.init_param_config()
        self.init_scanning_config()
        self.comboBox_3.currentIndexChanged.connect(self.init_param_config)
        self.comboBox.currentIndexChanged.connect(self.combox_change)
        self.lineEdit_2.returnPressed.connect(self.on_return_pressed)
        self.pushButton_2.clicked.connect(self.revise_quit_button_click)  # 刷新键
        self.pushButton.clicked.connect(self.revise_button_click)  # 保存修改
        self.pushButton_3.clicked.connect(self.recheck_button_click)

    def set_window_title(self):
        window_title = self.config_ini['STATION_ENV']['title'] + ' V2.3' + '    Line:' + \
                       self.config_ini['STATION_ENV']['linename'] + '    Station:' + \
                       self.config_ini['STATION_ENV']['stationno']
        self.setWindowTitle(window_title)

    def check_config_ini_path(self):
        if not os.path.exists(os.path.join(os.getcwd(), 'Config.ini')):
            error_message = "Config文件路径不存在"
            QMessageBox.critical(self, "Error", error_message, QMessageBox.Ok)
            sys.exit()

    def get_model_ini_path(self):
        self.combox_index = int(self.config_ini['FILE_PATH']['ModelSelect'])
        relative_path = self.config_ini['FILE_PATH']['ModelFile'].replace('/', '\\')
        path = os.path.join(os.getcwd(), relative_path)
        # print(path)
        if not os.path.exists(path):
            error_message = "model文件路径不存在"
            QMessageBox.critical(self, "Error", error_message, QMessageBox.Ok)
            sys.exit()
        return path

    def clear_rows(self):
        rows_to_remove = []
        for row_idx in range(self.box_num * self.innerQty, self.tableWidget_2.rowCount()):
            internal_no_item = self.tableWidget_2.item(row_idx, 0)
            if internal_no_item and internal_no_item.text() == str(self.box_num+1):
                rows_to_remove.append(row_idx)
        for row_idx in reversed(rows_to_remove):
            self.tableWidget_2.removeRow(row_idx)

    def init_combobox(self):
        self.comboBox.clear()
        self.comboBox_3.clear()
        for section in self.config.sections():
            self.comboBox.addItem(self.config[section]['model'].split(';')[0].strip())
            self.comboBox_3.addItem(self.config[section]['model'].split(';')[0].strip())
        self.comboBox.setCurrentIndex(self.combox_index)
        self.comboBox_3.setCurrentIndex(self.combox_index)

    def init_tablewidget(self):
        self.tableWidget_2.clearContents()
        self.tableWidget_2.setColumnCount(4)
        self.tableWidget_2.setHorizontalHeaderLabels(["内箱编号", "台次", "序列号", "结果"])
        widths = [100, 60, 230, 60]
        for col, width in enumerate(widths):
            self.tableWidget_2.setColumnWidth(col, width)
        self.tableWidget_2.horizontalHeader().setStretchLastSection(True)
        self.tableWidget_2.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)

    def init_param_config(self):  # 初始化参数配置
        self.pre_sncode = []
        model_name = self.comboBox_3.currentText()
        for section in self.config.sections():
            if model_name in self.config[section]['model']:
                self.model_param = section
        self.lineEdit_6.setText(self.config[self.model_param]['ProductName'].split(';')[0].strip())
        self.lineEdit_9.setText(self.config[self.model_param]['OrderNo'].split(';')[0].strip())
        self.lineEdit_10.setText(self.config[self.model_param]['InternalNo'].split(';')[0].strip())
        self.lineEdit_7.setText(self.config[self.model_param]['LotNo'].split(';')[0].strip())
        self.lineEdit_8.setText(self.config[self.model_param]['OuterBoxNo'].split(';')[0].strip())
        self.lineEdit_11.setText(self.config[self.model_param]['InnerQty'].split(';')[0].strip())
        self.lineEdit_12.setText(self.config[self.model_param]['OuterQty'].split(';')[0].strip())
        self.lineEdit_13.setText(self.config[self.model_param]['RatedCurrent'].split(';')[0].strip())
        self.lineEdit_14.setText(self.config[self.model_param]['TotalQty'].split(';')[0].strip())
        self.label_86.setText(self.config[self.model_param]['ChkDataPath'].split(';')[0].strip())
        self.label_88.setText(self.config[self.model_param]['SNDataPath'].split(';')[0].strip())
        self.label_90.setText(self.config[self.model_param]['PackDataPath'].split(';')[0].strip())
        output_name = self.config[self.model_param]['OrderNo'].split(';')[0].strip() + '_' + \
                      self.config[self.model_param]['InternalNo'].split(';')[0].strip() + '.PCK'
        self.label_92.setText(output_name)

    def init_scanning_config(self): # 初始化扫码包装
        self.pre_sncode = []
        model_name = self.comboBox.currentText()
        for section in self.config.sections():
            if model_name in self.config[section]['model']:
                self.model_scan = section
        self.checkBox.setEnabled(True)
        self.innerQty = int(self.config[self.model_scan]['InnerQty'].split(';')[0].strip())
        self.label_8.setText(self.config[self.model_scan]['OrderNo'].split(';')[0].strip())
        self.label_11.setText(self.config[self.model_scan]['InternalNo'].split(';')[0].strip())
        self.label_13.setText(self.config[self.model_scan]['LotNo'].split(';')[0].strip())
        self.label_14.setText(self.config[self.model_scan]['outerboxno'].split(';')[0].strip())
        self.label_16.setText(self.config[self.model_scan]['OuterQty'].split(';')[0].strip())
        self.label_18.setText(self.config[self.model_scan]['InnerQty'].split(';')[0].strip())
        self.label_20.setText(self.config[self.model_scan]['TotalQty'].split(';')[0].strip())
        self.total_check = 0 # 累计集合包装完成数量
        self.box_num = 0 # 内箱编号
        self.platform_num = 0 # 内箱台次
        self.check_num = 0
        # self.label_26.setText(f'{self.total_check}')
        self.label_4.setText(f'{self.box_num}')
        self.label_29.setText(f'{self.platform_num}')
        self.label_31.setText('Ready')
        self.label_31.setStyleSheet("background-color: white;")
        self.init_package_labels()
        self.init_tablewidget()

    def init_package_labels(self):
        idx = 1
        for label in self.labels:
            label.setStyleSheet("background-color: gray;")
            label.setText(f"{idx}")
            idx += 1
        self.label_98.setStyleSheet("background-color: orange;")

    def sncode_to_regex(self, sncode):  # 将sncode变为对应的正则表达式
        regex = ""
        for char in sncode:
            if char == 'x':
                regex += r'[0-9a-zA-Z]'
            elif char == 'n':
                regex += r'[0-9]'
            elif char == 'c':
                regex += r'[a-zA-Z]'
            else:
                regex += re.escape(char)
        return f"{regex}"

    def save_file_title(self, text_row):
        error_code = 0
        self.save_path = os.path.join(self.config[self.model_param]['PackDataPath'].split(';')[0].strip(), self.label_92.text())
        if not os.path.exists(self.config[self.model_param]['PackDataPath'].split(';')[0].strip()):
            error_message = "保存文件路径不存在"
            QMessageBox.critical(self, "Error", error_message, QMessageBox.Ok)
            error_code = 1
        if error_code:
            return -1
        else:
            if self.checkBox.isChecked():
                with open(self.save_path, 'w', encoding='utf-8') as file:
                    file.write(f"Product Model: {self.config[self.model_scan]['model'].split(';')[0].strip()}\t\tLOTNo.: "
                               f"{self.config[self.model_scan]['lotno'].split(';')[0].strip()}\t\tOutBoxNo: {self.config[self.model_scan]['outerboxno'].split(';')[0].strip()}\n")
                    file.write(f"InnerQTY: {self.config[self.model_scan]['innerqty'].split(';')[0].strip()}\t\t\t\tOuterQTY: "
                               f"{self.config[self.model_scan]['outerqty'].split(';')[0].strip()}\t\tTotalLotQTY: {self.config[self.model_scan]['totalqty'].split(';')[0].strip()}\n")
                    file.write(f"OderNo: {self.config[self.model_scan]['orderno'].split(';')[0].strip()}\n")
                    file.write("-------------------------------------------------------------------"
                               "-----------------------------------------------------------\n")
                    file.write("InnerNo\tNo.\tSerial No\t\tResult\t")
                    for i in range(5, 5+int(self.config[self.model_scan]['inspectitems'].split(';')[0].strip())):
                        file.write(f"{text_row[i]}\t")
                    file.write("Datatime\n")
                    file.write("-------------------------------------------------------------------"
                               "-----------------------------------------------------------\n")

    def on_return_pressed(self):
        self.checkBox.setEnabled(False)
        text = self.lineEdit_2.text()
        text_row = []
        pattern = self.sncode_to_regex(self.config[self.model_scan]['sncode'].split(';')[0].strip())
        # print(pattern)
        # print(text, len(text), re.match(pattern, text))
        if len(text) == 18 and re.match(pattern, text):
            flag, error_code = 0,0
            # print('match success')
            idcheck_path = os.path.join(self.config[self.model_scan]['sndatapath'].split(';')[0].strip(),
                                        self.config[self.model_scan]['InternalNo'].split(';')[0].strip())
            check_path = os.path.join(self.config[self.model_scan]['chkdatapath'].split(';')[0].strip(),
                                      self.config[self.model_scan]['InternalNo'].split(';')[0].strip())
            if not os.path.exists(idcheck_path):
                error_message = "序列号文件检查路径不存在"
                QMessageBox.critical(self, "Error", error_message, QMessageBox.Ok)
                error_code = 1
            if not os.path.exists(check_path):
                error_message = "检验数据检查路径不存在"
                QMessageBox.critical(self, "Error", error_message, QMessageBox.Ok)
                error_code = 1
            if error_code:
                return
            for sn_name in os.listdir(idcheck_path):
                last_six_num = text[-6:]
                if last_six_num in sn_name:
                    with open(os.path.join(idcheck_path,sn_name),'r') as idcheck_file:
                        content = idcheck_file.read()
                        date_pattern = r'(\d{4}-\d{2}-\d{2})'
                        match = re.search(date_pattern, content)
                        if match:
                            extracted_date = match.group(1).replace('-','')  # 获取捕获的年月日部分
                        else:
                            print("No valid date found in the file.")
                        for csv_name in os.listdir(check_path):
                            if extracted_date in csv_name:
                                with open(os.path.join(check_path, csv_name), 'r') as check_file:
                                    csv_reader = csv.reader(check_file)
                                    first_row = next(csv_reader)
                                    for row in csv_reader:
                                        if text in row:
                                            flag = 1
                                            text_row = row
                                            if not self.save_title_bool:
                                                self.save_file_title(first_row)
                                                self.save_title_bool = True
                                            break
            if self.check_num >= self.innerQty:
                self.box_num += 1
                self.init_package_labels()
                self.check_num = 0
                self.platform_num = 0
            if text in self.pre_sncode:
                self.labels[self.check_num].setStyleSheet("background-color: red;")
                self.label_31.setText('NG')
                self.label_31.setStyleSheet("background-color: red;")
                error_message = "输入了一个重复的PSN码"
                QMessageBox.critical(self, "Error", error_message, QMessageBox.Ok)
            elif flag == 0:
                self.labels[self.check_num].setStyleSheet("background-color: red;")
                self.label_31.setText('NG')
                self.label_31.setStyleSheet("background-color: red;")
                error_message = "找不到对应的PSN码"
                QMessageBox.critical(self, "Error", error_message, QMessageBox.Ok)
            else:
                self.pre_sncode.append(text)
                self.labels[self.check_num].setStyleSheet("background-color: limegreen;")
                self.label_31.setText('OK')
                self.label_31.setStyleSheet("background-color: limegreen;")
                if self.checkBox.isChecked():
                    with open(self.save_path, 'a', encoding='utf-8') as file:
                        line = f"{self.box_num + 1}\t{self.platform_num + 1}\t{text}\tOK\t"
                        for i in range(5, 5+int(self.config[self.model_scan]['inspectitems'].split(';')[0].strip())):
                            if not text_row[i]:
                                text_row[i] = "NA"
                            line += f"{text_row[i]}\t\t"
                        line += f"{str(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))}\n"
                        file.write(line)
                row = (f"{self.box_num + 1}", f"{self.platform_num + 1}", f"{text}", "OK")
                row_idx = self.tableWidget_2.rowCount()
                self.tableWidget_2.insertRow(row_idx)
                for col_idx, value in enumerate(row):
                    item = QTableWidgetItem(value)
                    color = QColor(50, 205, 50)
                    item.setBackground(color)
                    self.tableWidget_2.setItem(row_idx, col_idx, item)
                # self.tableWidget_2.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
                self.check_num += 1
                self.total_check += 1
                self.platform_num += 1
            if self.check_num < self.innerQty:
                self.labels[self.check_num].setStyleSheet("background-color: orange;")
        else:
            self.labels[self.check_num].setStyleSheet("background-color: red;")
            self.label_31.setText('NG')
            self.label_31.setStyleSheet("background-color: red;")
            error_message = "输入的PSN码错误"
            QMessageBox.critical(self, "Error", error_message, QMessageBox.Ok)
        self.label_4.setText(f'{self.box_num+1}')
        self.label_29.setText(f'{self.platform_num}')

    def revise_quit_button_click(self):
        self.lineEdit_6.setText(self.config[self.model_param]['ProductName'].split(';')[0].strip())
        self.lineEdit_9.setText(self.config[self.model_param]['OrderNo'].split(';')[0].strip())
        self.lineEdit_10.setText(self.config[self.model_param]['InternalNo'].split(';')[0].strip())
        self.lineEdit_7.setText(self.config[self.model_param]['LotNo'].split(';')[0].strip())
        self.lineEdit_8.setText(self.config[self.model_param]['OuterBoxNo'].split(';')[0].strip())
        self.lineEdit_11.setText(self.config[self.model_param]['InnerQty'].split(';')[0].strip())
        self.lineEdit_12.setText(self.config[self.model_param]['OuterQty'].split(';')[0].strip())
        self.lineEdit_13.setText(self.config[self.model_param]['RatedCurrent'].split(';')[0].strip())
        self.lineEdit_14.setText(self.config[self.model_param]['TotalQty'].split(';')[0].strip())

    def revise_button_click(self):
        self.config[self.model_param]['ProductName'] = self.lineEdit_6.text() + '\t\t'\
                                                       + ';' + self.config[self.model_param]['ProductName'].split(';')[1]
        self.config[self.model_param]['OrderNo'] = self.lineEdit_9.text() + '\t\t'\
                                                       + ';' + self.config[self.model_param]['OrderNo'].split(';')[1]
        self.config[self.model_param]['InternalNo'] = self.lineEdit_10.text() + '\t\t'\
                                                       + ';' + self.config[self.model_param]['InternalNo'].split(';')[1]
        self.config[self.model_param]['LotNo'] = self.lineEdit_7.text() + '\t\t'\
                                                       + ';' + self.config[self.model_param]['LotNo'].split(';')[1]
        self.config[self.model_param]['OuterBoxNo'] = self.lineEdit_8.text() + '\t\t'\
                                                       + ';' + self.config[self.model_param]['OuterBoxNo'].split(';')[1]
        self.config[self.model_param]['InnerQty'] = self.lineEdit_11.text() + '\t\t'\
                                                       + ';' + self.config[self.model_param]['InnerQty'].split(';')[1]
        self.config[self.model_param]['OuterQty'] = self.lineEdit_12.text() + '\t\t'\
                                                       + ';' + self.config[self.model_param]['OuterQty'].split(';')[1]
        self.config[self.model_param]['RatedCurrent'] = self.lineEdit_13.text() + '\t\t'\
                                                       + ';' + self.config[self.model_param]['RatedCurrent'].split(';')[1]
        self.config[self.model_param]['TotalQty'] = self.lineEdit_14.text() + '\t\t'\
                                                       + ';' + self.config[self.model_param]['TotalQty'].split(';')[1]
        print("保存成功")
        with open(self.model_ini_path, 'w', encoding="utf-8") as configfile:
            self.config.write(configfile)
        self.init_scanning_config()

    def remove_last_line(self, file_path):
        with open(file_path, "r+", encoding="utf-8") as file:
            file.seek(0, 2)  # 将文件指针移动到文件末尾
            # 定位到最后一行开始的位置
            pos = file.tell()
            pos -= 1
            while pos > 0:
                pos -= 1
                file.seek(pos)
                char = file.read(1)
                if char == "\n":
                    break
            # 截断文件到最后一行的位置
            # print(pos)
            file.truncate(pos)


    def recheck_button_click(self):
        for i in range(0, self.platform_num+1):
            if self.checkBox.isChecked():
                self.remove_last_line(self.save_path)
            if i < self.platform_num:
                self.pre_sncode.pop()
            if i == self.platform_num:
                with open(self.save_path, 'a', encoding='utf-8') as file:
                    file.write('\n')
        self.clear_rows()
        self.init_package_labels()
        self.check_num = 0
        self.platform_num = 0
        self.total_check = self.innerQty * self.box_num
        # self.label_26.setText(f'{self.total_check}')
        self.label_4.setText(f'{self.box_num}')
        self.label_29.setText(f'{self.platform_num}')

    def combox_change(self):  # 当combox改变时，让combox_3也改变，同时改变config中的参数
        self.init_scanning_config()
        self.combox_index = self.comboBox.currentIndex()
        self.comboBox_3.setCurrentIndex(self.combox_index)
        self.init_param_config()
        self.config_ini['FILE_PATH']['ModelSelect'] = str(self.combox_index)
        with open(os.path.join(os.getcwd(), 'Config.ini'), 'w', encoding="utf-8") as configfile:
            self.config_ini.write(configfile)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    MainWindow = Myui()
    MainWindow.setFixedSize(1200, 800)
    MainWindow.show()
    sys.exit(app.exec_())
