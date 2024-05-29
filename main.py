import sys
import datetime
import json
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QLineEdit, QMenuBar, \
    QAction, QMessageBox, QMainWindow, QMenu, QProgressBar, QTextEdit, QActionGroup, QSizePolicy, QGroupBox, QTableWidget, \
    QGridLayout, QTableWidget, QHeaderView, QAbstractItemView, QTableWidgetItem

import traceback
import qtmodern.styles
import qtmodern.windows
from PyQt5 import QtWidgets as qtw


def print_with_debug(msg):
    print(msg)
    traceback.print_exc()


def get_center_position(width, height):
    active_screen = QApplication.desktop().screenNumber(QApplication.desktop().cursor().pos())
    screen_geometry = QApplication.desktop().screenGeometry(active_screen)
    x = int((screen_geometry.width() - width) / 2 + screen_geometry.left())
    y = int((screen_geometry.height() - height) / 2 + screen_geometry.top())
    return x, y


def read_urls_from_file(filename):
    try:
        with open(filename, 'r') as file:
            data = json.load(file)
            return data.get('urls', [])
    except FileNotFoundError:
        write_urls_to_file(filename, [])
        return []
    except json.JSONDecodeError:
        return []


def write_urls_to_file(filename, urls):
    data = {'urls': urls}
    with open(filename, 'w') as file:
        json.dump(data, file)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.url_list = read_urls_from_file('urls.json')
        self.edit_box = QLineEdit("")
        self.edit_box.setPlaceholderText("신규 게시판 url 주소를 입력하세요")
        self.add_button = QPushButton("추가")
        self.delete_button = QPushButton("삭제")
        self.table = QTableWidget()
        self.url_list_widget = qtw.QListWidget(self)
        self.progress_bar = QProgressBar(self)
        self.log_edit_box = QTextEdit(self)
        self.init_ui()
        self.update_list()

    def init_ui(self):
        try:
            self.setWindowTitle('SmartWriter')
            self.edit_box.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            self.add_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
            self.delete_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

            inner_layout = QHBoxLayout()
            inner_layout .addWidget(self.edit_box, stretch=5)
            inner_layout .addWidget(self.add_button, stretch=1)
            inner_layout .addWidget(self.delete_button, stretch=1)

            group_box = QGroupBox()
            group_box.setStyleSheet("QGroupBox { border: 1px solid gray; }")
            group_box.setMaximumHeight(100)
            group_box.setLayout(inner_layout)

            self.url_list_widget.setSelectionMode(qtw.QAbstractItemView.MultiSelection)
            self.add_button.clicked.connect(self.on_add_button_click)
            self.delete_button.clicked.connect(self.on_delete_button_click)
            self.log_edit_box.setReadOnly(True)  # 읽기 전용으로 설정
            self.log_edit_box.setFixedHeight(200)  # 높이 설정

            main_layout = QVBoxLayout()
            main_layout.addWidget(group_box)
            main_layout.addWidget(self.url_list_widget)
            main_layout.addWidget(self.progress_bar)
            main_layout.addWidget(self.log_edit_box)

            container = QWidget()
            container.setLayout(main_layout)
            self.setCentralWidget(container)
            self.resize(800, 600)
            x, y = get_center_position(self.width(), self.height())
            self.move(x, y)
            self.add_log('프로그램 시작')
        except Exception as e:
            print_with_debug(e)

    def add_log(self, log):
        current_time = datetime.datetime.now()
        formatted_time = current_time.strftime("[%Y-%m-%d %H:%M:%S]")
        formatted_log = f"{formatted_time} {log}"
        print(formatted_log)
        self.log_edit_box.append(formatted_log)
        self.log_edit_box.verticalScrollBar().setValue(self.log_edit_box.verticalScrollBar().maximum())

    def on_add_button_click(self):
        url = self.edit_box.text()
        if url:
            self.url_list.append(url)
            write_urls_to_file('urls.json', self.url_list)
            self.update_list()
            self.edit_box.clear()

    def on_delete_button_click(self):
        indexes = self.url_list_widget.selectedIndexes()
        for index in reversed(indexes):
            self.url_list_widget.takeItem(index.row())
            del self.url_list[index.row()]
            write_urls_to_file('urls.json', self.url_list)

    def update_list(self):
        self.url_list_widget.clear()
        for url in self.url_list:
            self.url_list_widget.addItem(url)

    def get_selected_urls(self):
        urls = []
        indexes = self.url_list_widget.selectedIndexes()
        for index in indexes:
            urls.append(self.url_list[index.row()])
        return urls


if __name__ == "__main__":
    app = QApplication(sys.argv)
    qtmodern.styles.dark(app)
    main_win = MainWindow()
    main_win.show()
    sys.exit(app.exec_())