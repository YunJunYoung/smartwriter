import sys
import os
import datetime
import json
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QLineEdit, QMenuBar, \
    QAction, QMessageBox, QMainWindow, QMenu, QProgressBar, QTextEdit, QActionGroup, QSizePolicy, QGroupBox, QTableWidget, \
    QGridLayout, QTableWidget, QHeaderView, QAbstractItemView, QTableWidgetItem
from PyQt5.QtGui import QIcon

import traceback
import qtmodern.styles
import qtmodern.windows
from PyQt5 import QtWidgets as qtw
from PyQt5.QtCore import Qt

from selenium import webdriver
from selenium.webdriver.common.by import By
import requests
import sys
import time
from twocaptcha.solver import TwoCaptcha


api_key = '37f1af7f0d286cd9ad65892446c64ab7'
solver = TwoCaptcha(api_key, defaultTimeout=30, pollingInterval=5)


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


def read_data_from_file(filename):
    try:
        with open(filename, 'r') as file:
            data = json.load(file)
            return data.get('title', ''), data.get('content', '')
    except (FileNotFoundError, json.JSONDecodeError):
        return '', ''


def write_data_to_file(filename, title, content):
    data = {'title': title, 'content': content}
    with open(filename, 'w') as file:
        json.dump(data, file)


def read_api_count(filename='api_count.json'):
    if not os.path.exists(filename):
        write_api_count(0, filename)
    with open(filename, 'r') as file:
        data = json.load(file)
        return data.get('count', 0)


def write_api_count(count, filename='api_count.json'):
    data = {'count': count}
    with open(filename, 'w') as file:
        json.dump(data, file)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.write_button = QPushButton("글쓰기 시작")
        self.url_list = read_urls_from_file('urls.json')
        self.api_count = read_api_count()
        self.title, self.content = read_data_from_file('data.json')
        self.url_edit_box = QLineEdit("")
        self.url_edit_box.setPlaceholderText("신규 게시판 url 주소를 입력하세요")
        self.title_modify_button = QPushButton("변경")
        self.title_edit_box = QLineEdit(self.title)
        self.content_edit_box = QTextEdit(self.content)
        self.add_button = QPushButton("추가")
        self.delete_button = QPushButton("삭제")
        self.table = QTableWidget()
        self.url_list_widget = qtw.QListWidget(self)
        self.progress_bar = QProgressBar(self)
        self.log_edit_box = QTextEdit(self)
        self.count_label = None
        self.init_ui()
        self.update_list()
        self.driver = None

    def init_ui(self):
        try:
            self.setWindowTitle('SmartWriter')
            self.setWindowIcon(QIcon('main_icon.png'))
            self.url_edit_box.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            self.add_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
            self.delete_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

            inner_layout = QVBoxLayout()
            inner_layout_top_layout = QHBoxLayout()
            inner_layout_top_layout.addWidget(self.url_edit_box, stretch=5)
            inner_layout_top_layout.addWidget(self.add_button, stretch=1)
            inner_layout_top_layout.addWidget(self.delete_button, stretch=1)

            inner_layout.addLayout(inner_layout_top_layout)
            inner_layout.addWidget(self.url_list_widget)

            group_box = QGroupBox()
            group_box.setStyleSheet("QGroupBox { border: 1px solid gray; }")
            group_box.setLayout(inner_layout)

            self.url_list_widget.setSelectionMode(qtw.QAbstractItemView.MultiSelection)
            self.url_list_widget.setMaximumHeight(500)
            self.url_list_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # 크기 정책 설정
            self.add_button.clicked.connect(self.on_add_button_click)
            self.delete_button.clicked.connect(self.on_delete_button_click)

            main_layout = QVBoxLayout()
            main_layout.addWidget(self.write_button)

            self.count_label = QLabel(f"누적 캡챠 호출 횟수 : {self.api_count}", self)
            self.count_label.setAlignment(Qt.AlignCenter)
            main_layout.addWidget(self.count_label)

            self.write_button.clicked.connect(self.on_write_button_click)
            main_layout.addWidget(group_box)

            inner_layout_bottom = QVBoxLayout()
            inner_layout_bottom.addWidget(QLabel('제목'))
            inner_layout_bottom_title_layout = QHBoxLayout()
            inner_layout_bottom_title_layout.addWidget(self.title_edit_box)
            inner_layout_bottom_title_layout.addWidget(self.title_modify_button)

            self.title_modify_button.clicked.connect(self.on_title_modify_button_click)

            inner_layout_bottom.addLayout(inner_layout_bottom_title_layout)
            inner_layout_bottom.addWidget(QLabel('본문'))

            self.content_edit_box.setFixedHeight(150)
            inner_layout_bottom.addWidget(self.content_edit_box)

            group_box_bottom = QGroupBox()
            group_box_bottom.setStyleSheet("QGroupBox { border: 1px solid gray; }")
            group_box_bottom.setMaximumHeight(300)
            group_box_bottom.setLayout(inner_layout_bottom)
            main_layout.addWidget(group_box_bottom)

            main_layout.addWidget(self.progress_bar)

            self.log_edit_box.setReadOnly(True)  # 읽기 전용으로 설정
            self.log_edit_box.setFixedHeight(150)  # 높이 설정
            main_layout.addWidget(self.log_edit_box)

            container = QWidget()
            container.setLayout(main_layout)
            self.setCentralWidget(container)
            self.resize(800, 800)
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

    def init_web_driver(self):
        try:
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")  # 브라우저 최대화 옵션 추가
            options.add_argument(
                "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
            self.driver = webdriver.Chrome(options=options)
        except Exception as e:
            print_with_debug(e)

    def on_write_button_click(self):
        try:
            self.init_web_driver()
            urls = self.get_all_urls()

            for url in urls:
                try:
                    self.driver.get(url)
                    self.add_log(f'[{url}] 사이트 이동')
                    self.driver.implicitly_wait(3)
                    name_box = self.driver.find_element(By.CSS_SELECTOR, "#wr_name")
                    name_box.send_keys("홍길동")

                    phone_box = self.driver.find_element(By.CSS_SELECTOR, "#wr_homepage")
                    phone_box.send_keys("010-1234-5678")

                    password_box = self.driver.find_element(By.CSS_SELECTOR, "#wr_password")
                    password_box.send_keys("password1")

                    subject_box = self.driver.find_element(By.CSS_SELECTOR, "#wr_subject")
                    subject_box.send_keys(self.title)

                    content_box = self.driver.find_element(By.CSS_SELECTOR, "#wr_content")
                    content_box.send_keys(self.content)

                    email_box = self.driver.find_element(By.CSS_SELECTOR, "#wr_email")
                    email_box.send_keys("test@naver.com")

                    self.add_log(f'[{url}] 정보입력 완료')

                    time.sleep(5)
                    self.write_contents(url)
                    time.sleep(5)
                except Exception as e:
                    print_with_debug(e)

        except Exception as e:
            print_with_debug(e)

    def write_contents(self, url):
        try:
            self.driver.implicitly_wait(10)

            # 캡차 이미지 요소 찾기
            captcha_img = self.driver.find_element(By.CSS_SELECTOR, "#captcha_img")
            captcha_url = captcha_img.get_attribute("src")
            cookies = self.driver.get_cookies()

            # requests 세션 생성
            session = requests.Session()

            # Selenium 쿠키를 requests 쿠키로 복사
            for cookie in cookies:
                session.cookies.set(cookie['name'], cookie['value'])

            response = session.get(captcha_url, headers={
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            })

            if response.status_code == 200:

                # 파일을 바이너리 모드로 열고 저장
                with open("captcha.jpg", "wb") as file:
                    file.write(response.content)
                try:
                    result = solver.normal('captcha.jpg')
                except Exception as e:
                    sys.exit(e)
                else:
                    self.api_count += 1
                    self.count_label.setText(f"누적 캡챠 호출 횟수 : {self.api_count}")
                    write_api_count(self.api_count)
                    number = result['code']
                    self.add_log(f'[{url}] 캡챠 인식 완료 : {number}')

                    captcha_box = self.driver.find_element(By.CSS_SELECTOR, "#captcha_key")
                    captcha_box.send_keys(number)

                    # btn_submit
                    submit_button = self.driver.find_element(By.CSS_SELECTOR, "#btn_submit")
                    submit_button.click()
                    self.add_log(f'[{url}] 글쓰기 작성 완료')

        except Exception as e:
            print_with_debug(e)

    def on_add_button_click(self):
        url = self.url_edit_box.text()
        if url:
            self.url_list.append(url)
            write_urls_to_file('urls.json', self.url_list)
            self.update_list()
            self.url_edit_box.clear()

    def on_delete_button_click(self):
        indexes = self.url_list_widget.selectedIndexes()
        for index in reversed(indexes):
            self.url_list_widget.takeItem(index.row())
            del self.url_list[index.row()]
            write_urls_to_file('urls.json', self.url_list)

    def on_title_modify_button_click(self):
        title = self.title_edit_box.text()
        content = self.content_edit_box.toPlainText()
        write_data_to_file('data.json', title, content)
        self.add_log('데이터 저장 완료')

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

    def get_all_urls(self):
        urls = []
        for index in range(self.url_list_widget.count()):
            item = self.url_list_widget.item(index)
            urls.append(item.text())
        return urls


if __name__ == "__main__":
    app = QApplication(sys.argv)
    qtmodern.styles.dark(app)
    main_win = MainWindow()
    main_win.show()
    sys.exit(app.exec_())