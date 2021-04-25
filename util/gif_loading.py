# encoding: utf-8
"""
@author: zeng zonghai
@software: PyCharm
@file: catalog.py
@time: 2021/4/25 20:21
"""
from PyQt5.Qt import *
import time
import sys, os


class GifLoading(QMainWindow):
    def __init__(self, parent, gif=None, min=0):
        """
        :param parent:
        :param gif:
        :param min:
        """
        super(GifLoading, self).__init__(parent)

        self.min = min
        self.show_time = 0

        parent.installEventFilter(self)

        self.label = QLabel()
        
        if not gif is None:
            self.movie = QMovie(gif)
            self.label.setMovie(self.movie)
            self.label.setAttribute(Qt.WA_TranslucentBackground)
            self.label.setFixedSize(QSize(160, 160))
            self.label.setScaledContents(True)
            self.movie.start()

        layout = QHBoxLayout()
        widget = QWidget()
        widget.setObjectName('background')
        widget.setStyleSheet('QWidget#background{background-color: rgba(255, 255, 255, 40%);}')
        widget.setLayout(layout)
        layout.addWidget(self.label)

        self.setCentralWidget(widget)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Tool)
        self.hide()

    def eventFilter(self, widget, event):
        events = {QMoveEvent, QResizeEvent, QPaintEvent}
        if widget == self.parent():
            if type(event) == QCloseEvent:
                self.close()
                return True
            elif type(event) in events:
                self.moveWithParent()
                return True
        return super(GifLoading, self).eventFilter(widget, event)

    def moveWithParent(self):
        if self.parent().isVisible():
            self.move(self.parent().geometry().x(), self.parent().geometry().y())
            self.setFixedSize(QSize(self.parent().geometry().width(), self.parent().geometry().height()))

    def show(self):
        super(GifLoading, self).show()
        self.show_time = time.time()
        self.moveWithParent()

    def close(self):
        # 显示时间不够最小显示时间 设置Timer延时删除
        if (time.time() - self.show_time) * 1000 < self.min:
            QTimer().singleShot((time.time() - self.show_time) * 1000 + 10, self.close)
        else:
            super(GifLoading, self).hide()
            super(GifLoading, self).deleteLater()


if __name__ == '__main__':
    app = QApplication(sys.argv)

    widget = QWidget()
    widget.setFixedSize(500, 500)
    widget.setStyleSheet('QWidget{background-color:white;}')

    button = QPushButton('button')
    layout = QHBoxLayout()
    layout.addWidget(button)
    widget.setLayout(layout)

    gif_path = os.getcwd() + '/../img/loading.gif'
    print(gif_path)
    loading_mask = GifLoading(widget, gif_path)
    widget.show()
    loading_mask.show()

    # QTimer().singleShot(1000, lambda: loading_mask.hide())

    sys.exit(app.exec_())
