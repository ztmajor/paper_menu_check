# encoding: utf-8
"""
@author: duke mei
@contact: mylovezju@163.com
@software: PyCharm
@file: main.py
@time: 2021/4/14 18:44
"""
import sys
import os

from PyQt5.QtWidgets import (QWidget, QToolTip, QDesktopWidget, QMessageBox, QTextEdit, QLabel,
                             QPushButton, QApplication, QMainWindow, QAction, qApp, QHBoxLayout, QVBoxLayout,
                             QGridLayout, QFileDialog, QLineEdit, QTextBrowser)
from PyQt5.QtGui import QFont, QIcon, QColor
from PyQt5.QtCore import QCoreApplication, pyqtSlot


from util import check_catalog, check_footer, gif_loading, transfer_thread


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.title = 'Check The List of Figures and Tables'
        self.left, self.top, self.width, self.height = 300, 300, 480, 270
        self.temp_docx_path = None
        self.temp_pdf_path = None
        self.loading = None
        self.back_file_transfer_thread = None
        self.initUI()

    def initUI(self):
        # 这里使用10px大小的SansSerif字体。
        QToolTip.setFont(QFont('SansSerif', 16))  # 这个静态方法设置了用于提示框的字体。

        self.setWindowTitle(self.title)
        # 窗口在屏幕上显示，并设置了它的尺寸。resize()和remove()合而为一的方法。
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.setWindowIcon(QIcon(os.getcwd() + '/pic/paper.png'))  # 创建一个QIcon对象并接收一个我们要显示的图片路径作为参数。

        # open button
        self.open_btn = QPushButton('Open', self)
        self.open_btn.setToolTip('Click <b>Open</b> to open the file')  # 调用setTooltip()方法创建提示框。
        self.open_btn.clicked.connect(self.open_slot_method)  # 点击事件
        # self.open_btn.resize(self.open_btn.sizeHint())  # 改变按钮大小
        # self.open_btn.move(50, 220)  # 移动按钮位置

        # test button
        self.check_btn = QPushButton('Check', self)
        self.check_btn.clicked.connect(self.check_slot_method)
        self.check_btn.setToolTip('Click <b>Check</b> to check the paper')
        # self.check_btn.resize(self.check_btn.sizeHint())
        # self.check_btn.move(150, 220)
        self.check_btn.setEnabled(False)

        # clean button
        self.clear_btn = QPushButton('Clear', self)
        self.clear_btn.clicked.connect(self.clear_slot_method)
        self.clear_btn.setToolTip('Click <b>Clear</b> to clear the textbox')

        self.textbox = QTextEdit(self)
        self.textbox.setReadOnly(True)
        self.textbox.document().setMaximumBlockCount(100)
        # self.textbox.move(50, 50)
        # self.textbox.resize(380, 150)

        self.create_grid_layout()
        self.show()

    def open_file_name_dialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self, "open doc file", "",
                                                  "Doc Files (*.doc)", options=options)
        if fileName:
            print(fileName)
        return fileName

    @pyqtSlot()
    def open_slot_method(self):
        print('open method called.')
        doc_file = self.open_file_name_dialog()
        temp_path = os.getcwd() + "/temp/"
        if not os.path.exists(temp_path):
            os.mkdir(temp_path)

        temp_path = temp_path + doc_file.split('/')[-1]
        self.temp_docx_path = temp_path.replace('.doc', '.docx')
        self.temp_pdf_path = temp_path.replace('.doc', '.pdf')
        print(self.temp_docx_path, self.temp_pdf_path)
        self.print_log(f'打开文件： {doc_file}', 'blue')
        self.print_log('检查是否需要转换成易检查的格式...')
        self.check_btn.setEnabled(False)
        if not os.path.exists(self.temp_docx_path) or not os.path.exists(self.temp_pdf_path):
            self.print_log('第一次检查该文件，需要进行转换，请稍等片刻...', 'blue')
            QApplication.processEvents()

            self.back_file_transfer_thread = transfer_thread.BackFileTransferQthread()
            self.back_file_transfer_thread.set_path(doc_file, self.temp_docx_path, self.temp_pdf_path)
            self.back_file_transfer_thread.transfer_error.connect(self.handle_file_transfer_error)
            self.back_file_transfer_thread.task_done.connect(self.handle_file_transfer_done)

            self.loading = gif_loading.GifLoading(self, os.getcwd() + '/img/loading.gif')
            self.loading.show()

            self.back_file_transfer_thread.start()
        else:
            self.transfer_done()

    def handle_file_transfer_error(self, err_log):
        self.loading.close()
        self.print_log(err_log, 'red')

    def handle_file_transfer_done(self):
        self.loading.close()
        self.transfer_done()

    def transfer_done(self):
        self.print_log('文件转换成功，可以开始检查。', 'green')
        self.check_btn.setEnabled(True)

    @pyqtSlot()
    def check_slot_method(self):
        print('click test')
        reply = QMessageBox.question(self, 'Message',
                                     "Start check?",
                                     QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
                                     QMessageBox.Yes)
        if reply == QMessageBox.Yes:
            print('let us check!')
            self.check('图')
            self.check('表')

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
            self.check_btn.setEnabled(False)
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

    def check(self, catalog_prefix='图'):
        self.print_log(f"开始检查{catalog_prefix}目录...")
        catalog_msg_list = check_catalog(self.temp_pdf_path, catalog_prefix)
        footer_msg_list = check_footer(self.temp_pdf_path, self.temp_docx_path)
        msg_list = catalog_msg_list + footer_msg_list
        # msg_list = catalog_msg_list
        if len(msg_list) == 0:
            self.print_log(f"{catalog_prefix}目录检查完毕，未发现错误。\n", 'green')
        else:
            for msg in msg_list:
                self.print_log(str(msg), 'red')
            self.print_log(f"{catalog_prefix}目录检查完毕。\n")

    def create_grid_layout(self):
        grid = QGridLayout()
        grid.setSpacing(10)  # 创建了一个网格布局并且设置了组件之间的间距

        grid.addWidget(self.textbox, 0, 0, 5, 3)  # 如果我们向网格布局中增加一个组件，我们可以提供组件的跨行
        grid.addWidget(self.open_btn, 6, 0)
        grid.addWidget(self.check_btn, 6, 1)
        grid.addWidget(self.clear_btn, 6, 2)

        self.setLayout(grid)

    def print_log(self, massage, color='black'):
        self.textbox.setTextColor(QColor(color))
        self.textbox.append(massage)
        print(massage)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MainWindow()
    sys.exit(app.exec_())
