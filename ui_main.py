# encoding: utf-8
"""
@author: duke mei
@contact: mylovezju@163.com
@software: PyCharm
@file: ui_main.py
@time: 2021/4/16 8:53
"""
import sys
import os
from PyQt5.QtWidgets import (QWidget, QToolTip, QDesktopWidget, QMessageBox, QTextEdit, QLabel,
                             QPushButton, QApplication, QMainWindow, QAction, qApp, QHBoxLayout, QVBoxLayout,
                             QGridLayout, QFileDialog, QLineEdit, QTextBrowser)
from PyQt5.QtGui import QFont, QIcon
from PyQt5.QtCore import QCoreApplication, pyqtSlot

from util import *


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.title = 'Check The List of Figures and Tables'
        self.left, self.top, self.width, self.height = 300, 300, 480, 270
        self.temp_docx_path = None
        self.temp_pdf_path = None
        self.initUI()

    def initUI(self):
        # 这里使用10px大小的SansSerif字体。
        QToolTip.setFont(QFont('SansSerif', 16))  # 这个静态方法设置了用于提示框的字体。

        self.setWindowTitle(self.title)
        # 窗口在屏幕上显示，并设置了它的尺寸。resize()和remove()合而为一的方法。
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.setWindowIcon(QIcon(r'W:\ZJU\课程\春学期\写作指导\作业\hw\pic\zju_icon.jpg'))  # 创建一个QIcon对象并接收一个我们要显示的图片路径作为参数。

        # open button
        self.open_btn = QPushButton('Open', self)
        self.open_btn.setToolTip('Click <b>Open</b> to open the file')  # 调用setTooltip()方法创建提示框。
        self.open_btn.clicked.connect(self.open_slot_method)  # 点击事件
        # self.open_btn.resize(self.open_btn.sizeHint())  # 改变按钮大小
        # self.open_btn.move(50, 220)  # 移动按钮位置

        # test button
        self.check_btn = QPushButton('Check', self)
        self.check_btn.clicked.connect(self.check_slot_method)
        # self.check_btn.resize(self.check_btn.sizeHint())
        # self.check_btn.move(150, 220)
        self.check_btn.setEnabled(False)

        # clean button
        self.clear_btn = QPushButton('Clear', self)
        self.clear_btn.clicked.connect(self.clear_slot_method)

        self.textbox = QTextEdit(self)
        self.textbox.setReadOnly(True)
        # self.textbox.move(50, 50)
        # self.textbox.resize(380, 150)

        self.create_grid_layout()
        self.show()

    def open_file_name_dialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self, "QFileDialog.getOpenFileName()", "",
                                                  "Doc Files (*.doc)", options=options)
        if fileName:
            print(fileName)
        return fileName

    @pyqtSlot()
    def open_slot_method(self):
        print('open method called.')
        doc_file = self.open_file_name_dialog()
        self.temp_docx_path = doc_file.replace('.doc', '.docx')
        self.temp_pdf_path = doc_file.replace('.doc', '.pdf')
        self.textbox.append(f'Open {doc_file}')
        self.print_log('processing...')
        try:
            doc2docx(doc_file, self.temp_docx_path)
            doc2pdf(doc_file, self.temp_pdf_path)
            self.print_log('processing complete.')
            self.check_btn.setEnabled(True)
        except:
            self.print_log('process file error!')
            pass


    @pyqtSlot()
    def check_slot_method(self):
        print('click test')
        reply = QMessageBox.question(self, 'Message',
                                     "Start check?",
                                     QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
                                     QMessageBox.Yes)
        if reply == QMessageBox.Yes:
            print('let us check!')

            # TODO: doing sth
            msg_list = find_page_number(self.temp_pdf_path, '图 2')
            for msg in msg_list:
                self.print_log(str(msg))
            self.print_log('check complete.')

            os.remove(self.temp_docx_path)
            os.remove(self.temp_pdf_path)
        elif reply == QMessageBox.No:
            print('do not check')
        else:
            print('cancel check')

    @pyqtSlot()
    def clear_slot_method(self):
        print('clean method called.')
        reply = QMessageBox.question(self, 'Message',
                                     "Are you sure to clear the textbox?",
                                     QMessageBox.Yes | QMessageBox.No,
                                     QMessageBox.Yes)
        if reply == QMessageBox.Yes:
            print('clear textbox')
            self.textbox.clear()
        elif reply == QMessageBox.No:
            print('do not want to clear textbox')

    def closeEvent(self, event):
        """关闭时会弹出提示框"""
        reply = QMessageBox.question(self, 'Message',
                                     "Are you sure to quit?",
                                     QMessageBox.Yes | QMessageBox.No,
                                     QMessageBox.No)
        if reply == QMessageBox.Yes:
            event.accept()
        else:
            print('Pretend to close :)')
            event.ignore()

    def create_grid_layout(self):
        grid = QGridLayout()
        grid.setSpacing(10)  # 创建了一个网格布局并且设置了组件之间的间距

        grid.addWidget(self.textbox, 0, 0, 5, 3)  # 如果我们向网格布局中增加一个组件，我们可以提供组件的跨行
        grid.addWidget(self.open_btn, 6, 0)
        grid.addWidget(self.check_btn, 6, 1)
        grid.addWidget(self.clear_btn, 6, 2)

        self.setLayout(grid)

    def print_log(self, massage):
        print(massage)
        self.textbox.append(massage)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MainWindow()
    sys.exit(app.exec_())
