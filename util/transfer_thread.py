# encoding: utf-8
"""
@author: zeng zonghai
@software: PyCharm
@file: catalog.py
@time: 2021/4/25 20:38
"""
from PyQt5.Qt import *
import time
import sys
import os
from util import doc2pdf, doc2docx


class BackFileTransferQthread(QThread):
    transfer_error=pyqtSignal(str)
    task_done=pyqtSignal()

    def __init__(self, parent=None):
        super(BackFileTransferQthread, self).__init__(parent)
        self.doc_path = None
        self.docx_path = None
        self.pdf_path = None

    def set_path(self, temp_doc_path, temp_docx_path, temp_pdf_path):
        self.doc_path = temp_doc_path
        self.docx_path = temp_docx_path
        self.pdf_path = temp_pdf_path

    def run(self):
        try:
            if not os.path.exists(self.docx_path):
                doc2docx(self.doc_path, self.docx_path)
            if not os.path.exists(self.pdf_path):
                doc2pdf(self.doc_path, self.pdf_path)

            self.task_done.emit()
        except:
            self.transfer_error.emit('文件转换失败！')
