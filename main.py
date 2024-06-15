import sys
import os
import datetime
import json
import re
import urllib.parse
import clipboard

from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QLineEdit, QMenuBar, \
    QAction, QMessageBox, QMainWindow, QMenu, QProgressBar, QTextEdit, QActionGroup, QSizePolicy, QGroupBox, \
    QTableWidget, \
    QGridLayout, QTableWidget, QHeaderView, QAbstractItemView, QTableWidgetItem, QCheckBox, QListWidget, QFileDialog

from PyQt5.QtGui import QIcon
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.webdriver.common.action_chains import ActionChains

import traceback
import qtmodern.styles
import qtmodern.windows
from PyQt5 import QtWidgets as qtw
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from selenium.common.exceptions import TimeoutException

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import UnexpectedAlertPresentException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.service import Service
import requests
import time
from twocaptcha.solver import TwoCaptcha
import pandas as pd
import pyperclip
from selenium.webdriver.common.keys import Keys

import random

# 한글 초성, 중성, 종성 리스트
chosung = ['ㄱ', 'ㄲ', 'ㄴ', 'ㄷ', 'ㄸ', 'ㄹ', 'ㅁ', 'ㅂ', 'ㅃ', 'ㅅ', 'ㅆ', 'ㅇ', 'ㅈ', 'ㅉ', 'ㅊ', 'ㅋ', 'ㅌ', 'ㅍ', 'ㅎ']
jungsung = ['ㅏ', 'ㅐ', 'ㅑ', 'ㅒ', 'ㅓ', 'ㅔ', 'ㅕ', 'ㅖ', 'ㅗ', 'ㅘ', 'ㅙ', 'ㅚ', 'ㅛ', 'ㅜ', 'ㅝ', 'ㅞ', 'ㅟ', 'ㅠ', 'ㅡ', 'ㅢ', 'ㅣ']
jongsung = [''] + ['ㄱ', 'ㄲ', 'ㄳ', 'ㄴ', 'ㄵ', 'ㄶ', 'ㄷ', 'ㄹ', 'ㄺ', 'ㄻ', 'ㄼ', 'ㄽ', 'ㄾ', 'ㄿ', 'ㅀ', 'ㅁ', 'ㅂ', 'ㅄ', 'ㅅ', 'ㅆ',
                   'ㅇ', 'ㅈ', 'ㅊ', 'ㅋ', 'ㅌ', 'ㅍ', 'ㅎ']


def generate_korean_name():
    name_length = random.choice([3, 4])  # 이름 길이를 3 또는 4로 랜덤하게 선택
    name = ''

    for _ in range(name_length):
        ch = random.choice(chosung)
        ju = random.choice(jungsung)
        jo = random.choice(jongsung)
        # 한글 유니코드 생성
        name += chr(0xAC00 + (chosung.index(ch) * 588) + (jungsung.index(ju) * 28) + jongsung.index(jo))

    return name


def generate_korean_phone_number():
    # 010-으로 시작
    phone_number = "010-"

    # 중간 4자리 숫자 생성
    middle_four_digits = ''.join([str(random.randint(0, 9)) for _ in range(4)])

    # 마지막 4자리 숫자 생성
    last_four_digits = ''.join([str(random.randint(0, 9)) for _ in range(4)])

    # 전체 전화번호 조합
    phone_number += f"{middle_four_digits}-{last_four_digits}"

    return phone_number


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


def load_credentials_from_json(json_file):
    with open(json_file, 'r') as f:
        credentials = json.load(f)
    return credentials


def save_credentials_to_json(json_file, credentials):
    with open(json_file, 'w') as f:
        json.dump(credentials, f, indent=4)


def get_login_url(write_url):
    try:
        # 도메인 추출을 위한 정규식
        pattern = r"(https?://[^/]+)"
        # 도메인 추출
        match = re.match(pattern, write_url)
        if match:
            domain = match.group(1)
            login_url = f"{domain}/bbs/login.php"
            return login_url
        else:
            print("Invalid URL format.")
            return None
    except Exception as e:
        print_with_debug(e)
        return None


def read_excel_data(file_path):
    # 엑셀 파일을 읽어 DataFrame으로 변환
    df = pd.read_excel(file_path)

    # 첫 번째 행은 데이터가 아니므로 건너뜀
    data = df.iloc[1:]

    titles = data['제목'].tolist()  # 제목 열을 리스트로 변환
    contents = data['내용'].tolist()  # 내용 열을 리스트로 변환

    return titles, contents


def get_random_title_content(file_path):
    titles, contents = read_excel_data(file_path)

    if len(titles) != len(contents):
        raise ValueError("Titles and contents lists must be of the same length")

    random_index = random.randint(0, len(titles) - 1)

    return titles[random_index], contents[random_index]


# 엑셀 파일 경로
excel_file_path = "ad.xlsx"


# Example usage
def replace_spaces_with_decoded_unicode(text: str, unicode_file: str) -> str:
    # Load the unicode characters from the JSON file
    with open(unicode_file, 'r', encoding='utf-8') as json_file:
        unicode_list = json.load(json_file)

    decoded_unicode_list = []
    for char in unicode_list:
        try:
            # Try to decode the unicode character
            decoded_unicode_list.append(urllib.parse.unquote(char))
        except Exception as e:
            print(f"Error decoding {char}: {e}")

    # Convert the text by replacing spaces with random decoded unicode characters
    result = []
    for char in text:
        if char == ' ':
            while True:
                try:
                    replacement_char = random.choice(decoded_unicode_list)
                    # Verify that the chosen unicode character is valid by encoding and decoding it
                    if replacement_char.encode('utf-8').decode('utf-8') == replacement_char:
                        result.append(replacement_char)
                        break
                except Exception as e:
                    # If an error occurs, continue to choose another random character
                    print(f"Error with character {replacement_char}: {e}")
                    continue
        else:
            result.append(char)

    return ''.join(result)


def find_login_url(write_url):
    try:
        with open('login_urls.json', 'r') as f:
            data = json.load(f)
            for entry in data:
                if entry['write_url'] == write_url:
                    return entry['login_url']
    except FileNotFoundError:
        return None
    return None


class Worker(QThread):
    progress_updated = pyqtSignal(int)
    log_updated = pyqtSignal(str)
    api_count_updated = pyqtSignal(int)

    def __init__(self, entries, writing_delay, overall_delay, repeat, convert, excel_file_list):
        super().__init__()
        self.entries = entries
        self.api_count = read_api_count()
        self.writing_delay = writing_delay
        self.overall_delay = overall_delay
        self.repeat = repeat
        self.driver = None
        self.convert = convert
        self.excel_file_list = excel_file_list
        self.current_file_index = 0

    def init_web_driver(self):
        try:
            chrome_driver_path = ChromeDriverManager().install()
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")
            options.add_argument(
                f"user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36 Edg/91.0.864.59")
            self.driver = webdriver.Chrome(service=Service(chrome_driver_path), options=options)
            self.driver.implicitly_wait(3)
            self.driver.delete_all_cookies()
        except Exception as e:
            print_with_debug(e)

    def get_current_excel_file(self):
        try:
            current_file = self.excel_file_list[self.current_file_index]
            print(f'current_file : {current_file}')
            return current_file
        except Exception as e:
            print_with_debug(e)
            return None

    def set_value_with_javascript(self, element, text):
        try:
            pyperclip.copy(text)
            element.click()
            element.send_keys(Keys.CONTROL, 'v')
            time.sleep(1)
        except Exception as e:
            print_with_debug(e)

    def set_value_with_clipboard(self, element, text):
        clipboard.copy(text)
        element.click()
        element.send_keys(Keys.CONTROL, 'v')
        time.sleep(1)  # 클립보드에서 붙여넣기를 위한 대기 시간

    def switch_to_iframe_if_exists(self, selectors):
        iframe_exist = False
        for selector in selectors:
            try:
                iframe = self.driver.find_element(By.CSS_SELECTOR, selector)
                self.driver.switch_to.frame(iframe)
                iframe_exist = True
                print(f'Switched to iframe using selector: {selector}')
                break  # 첫 번째로 발견한 iframe으로 전환 후 종료
            except Exception as e:
                print(f'No iframe found with selector: {selector}')
                continue  # 다음 셀렉터로 이동
        return iframe_exist

    def starts_with_imweb(self, url):
        return url.startswith('[imweb]')

    def remove_url_prefix(self, url):
        if self.starts_with_imweb(url):
            return url[len('[imweb]'):]
        return url

    def _set_value_with_javascript(self, driver, element, text):
        try:
            # JavaScript를 사용하여 값을 설정하고 입력 이벤트를 트리거합니다.
            driver.execute_script("""
            arguments[0].value = arguments[1];
            arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            """, element, text)
            time.sleep(1)
        except Exception as e:
            print(f"Error: {e}")

    def set_contenteditable_value(self, driver, element, text):
        try:
            # JavaScript를 사용하여 값을 설정합니다.
            driver.execute_script("""
                arguments[0].innerHTML = arguments[1];
                arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                """, element, text)
            time.sleep(1)
        except Exception as e:
            print(f"Error: {e}")

    def run(self):
        for k in range(self.repeat):
            self.log_updated.emit(f'{k + 1}번째 반복 실행')
            for i, entry in enumerate(self.entries):
                try:
                    url, user_id, password = entry
                    self.init_web_driver()

                    if self.starts_with_imweb(url):
                        url = self.remove_url_prefix(url)
                        self.driver.get(url)
                        self.log_updated.emit(f'[Index:{k + 1}_{i + 1}] [{url}] 글쓰기 페이지 이동')
                        time.sleep(3)

                        for _ in range(4):
                            webdriver.ActionChains(self.driver).send_keys(Keys.TAB).perform()
                            time.sleep(0.5)  # 각 탭 사이에 약간의 지연을 추가하여 자연스럽게 보이도록 합니다

                        try:
                            # 요소가 로드될 때까지 대기 (최대 10초)
                            name_input_element = WebDriverWait(self.driver, 10).until(
                                EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="nick"]'))
                            )
                            print("Name input element found!")

                            # 요소가 화면에 보이도록 스크롤
                            self.driver.execute_script("arguments[0].scrollIntoView();", name_input_element)

                            # JavaScript를 사용하여 값을 설정
                            self._set_value_with_javascript(self.driver, name_input_element, 'ghfkddl')
                        except Exception as e:
                            print(f"Error: {e}")

                        try:
                            # 요소가 로드될 때까지 대기 (최대 10초)
                            password_input_element = WebDriverWait(self.driver, 10).until(
                                EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="secret_pass"]'))
                            )

                            # 요소가 화면에 보이도록 스크롤
                            self.driver.execute_script("arguments[0].scrollIntoView();", password_input_element)
                            self._set_value_with_javascript(self.driver, password_input_element, "1234%^&*")
                        except Exception as e:
                            print(f"Error: {e}")

                        title, content = get_random_title_content(self.get_current_excel_file())
                        try:
                            # 요소가 로드될 때까지 대기 (최대 10초)
                            subject_box = WebDriverWait(self.driver, 10).until(
                                EC.presence_of_element_located(
                                    (By.XPATH, '//*[@id="post_subject"]'))
                            )
                            subject_input_element = self.driver.find_element(By.ID, "post_subject")
                            self.driver.execute_script("arguments[0].scrollIntoView();", subject_input_element)
                            self._set_value_with_javascript(self.driver, subject_input_element, title)
                        except Exception as e:
                            print(f"Error: {e}")

                        title, content = get_random_title_content(self.get_current_excel_file())
                        try:
                            WebDriverWait(self.driver, 10).until(
                                EC.presence_of_element_located((By.ID, "post_body"))
                            )
                            # 실제 입력 요소 찾기
                            contenteditable_element = WebDriverWait(self.driver, 10).until(
                                EC.presence_of_element_located((By.CSS_SELECTOR, "#post_body .fr-element.fr-view"))
                            )
                            # 요소가 화면에 보이도록 스크롤
                            self.driver.execute_script("arguments[0].scrollIntoView();", contenteditable_element)

                            # JavaScript를 사용하여 값을 설정
                            self.set_contenteditable_value(self.driver, contenteditable_element,
                                                      content)
                        except Exception as e:
                            print(f"Error: {e}")

                        try:
                            # 버튼 찾기 및 클릭
                            submit_button = WebDriverWait(self.driver, 10).until(
                                EC.presence_of_element_located((By.CSS_SELECTOR, 'button._save_post.save_post.btn'))
                            )
                            self.driver.execute_script("arguments[0].scrollIntoView();", submit_button)
                            submit_button.click()
                        except Exception as e:
                            print(f'[{url}] 작성 버튼을 찾을 수 없습니다. 오류: {str(e)}')
                    else:
                        if user_id != '' and password != '':
                            login_need = True
                        else:
                            login_need = False

                        wait = WebDriverWait(self.driver, 3)  # 최대 20초 대기

                        if login_need:
                            login_url = find_login_url(url) or get_login_url(url)
                            self.log_updated.emit(f'[Index:{k + 1}_{i + 1}] [{login_url}] 로그인 페이지 이동')
                            self.driver.get(login_url)

                            try:
                                id_box = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#login_id")))
                                id_box.send_keys(user_id)
                            except Exception as e:
                                print_with_debug(e)

                            try:
                                # 명시적 대기를 사용하여 비밀번호 입력 상자가 로드될 때까지 대기
                                pw_box = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#login_pw")))
                                pw_box.send_keys(password)
                            except Exception as e:
                                print_with_debug(e)

                            # 시도할 로그인 버튼의 CSS 선택자 리스트
                            login_selectors = [
                                "#login_fs > input.btn_submit",
                                "#login_fs > button",
                                "#login_frm > input.btn_submit",
                                "#mb_login > form > div:nth-child(7) > button"
                            ]

                            login_button = None

                            for selector in login_selectors:
                                try:
                                    login_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                                    login_button.click()
                                    break  # 로그인 버튼을 찾고 클릭한 후 반복문 탈출
                                except TimeoutException:
                                    continue  # 현재 선택자로 로그인 버튼을 찾지 못한 경우 다음 선택자로 시도

                            time.sleep(1)

                        self.driver.get(url)
                        self.log_updated.emit(f'[Index:{k + 1}_{i + 1}] [{url}] 글쓰기 페이지 이동')

                        if login_need is False:
                            try:
                                name_box = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#wr_name")))
                                name_box.send_keys(generate_korean_name())
                            except Exception as e:
                                self.log_updated.emit(f'[{url}] 이름 입력 요소를 찾을 수 없습니다.')

                            try:
                                phone_box = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#wr_homepage")))
                                phone_box.send_keys(generate_korean_phone_number())
                            except Exception as e:
                                self.log_updated.emit(f'[{url}] 전화번호 입력 요소를 찾을 수 없습니다.')

                            try:
                                password_box = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#wr_password")))
                                password_box.send_keys("password1")
                            except Exception as e:
                                self.log_updated.emit(f'[{url}] 비밀번호 입력 요소를 찾을 수 없습니다.')

                            try:
                                email_box = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#wr_email")))
                                email_box.send_keys("test@naver.com")
                            except Exception as e:
                                self.log_updated.emit(f'[{url}] 이메일 입력 요소를 찾을 수 없습니다.')

                        #title, content = get_random_title_content(excel_file_path)
                        title, content = get_random_title_content(self.get_current_excel_file())
                        if self.convert:
                            title = replace_spaces_with_decoded_unicode(title, 'special_char.json')
                        try:
                            subject_box = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#wr_subject")))
                            self.set_value_with_javascript(subject_box, title)
                        except Exception as e:
                            print(f'[{url}] 제목 입력 요소를 찾을 수 없습니다. 오류: {str(e)}')

                        iframe_selectors = [
                            '#fwrite > ul > li:nth-child(4) > iframe',
                            '#fwrite > div:nth-child(16) > div > iframe'
                            # 필요에 따라 여기에 더 많은 셀렉터를 추가할 수 있습니다.
                        ]

                        # iframe이 있는지 확인하고 전환
                        iframe_exist = self.switch_to_iframe_if_exists(iframe_selectors)
                        if iframe_exist:
                            try:
                                html_selector = '#smart_editor2_content > div.se2_conversion_mode > ul > li:nth-child(2) > button'
                                self.driver.find_element(By.CSS_SELECTOR, html_selector).click()
                                pyperclip.copy(content)
                                ActionChains(self.driver).key_down(Keys.CONTROL).send_keys('v').key_up(
                                    Keys.CONTROL).perform()
                                self.driver.switch_to.default_content()
                            except Exception as e:
                                self.log_updated.emit(f'[{url}] iframe 내 본문/제목 요소를 찾을 수 없습니다.')
                        else:
                            try:
                                content_box = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#wr_content")))
                                self.driver.execute_script("arguments[0].scrollIntoView();", content_box)
                                self.set_value_with_clipboard(content_box, content)
                            except Exception as e:
                                print(f'[{url}] 본문 입력 요소를 찾을 수 없습니다. 오류: {str(e)}')

                        self.write_contents(url, login_need)

                    time.sleep(5)
                    self.log_updated.emit(f'[Index:{k + 1}_{i + 1}] [{url}] 글쓰기 완료')
                    self.progress_updated.emit(i + 1)
                    self.driver.quit()
                    if i + 1 != len(self.entries):
                        self.log_updated.emit(f'다음 글쓰기 까지 {self.writing_delay}초 대기')
                        time.sleep(self.writing_delay)
                except UnexpectedAlertPresentException as e:
                    self.driver.quit()
                    self.init_web_driver()
                except Exception as e:
                    print_with_debug(e)
            self.log_updated.emit(f'다음 반복 까지 {self.overall_delay}초 대기')
            self.current_file_index = (self.current_file_index + 1) % len(self.excel_file_list)
            time.sleep(self.overall_delay)
        self.driver.quit()

    def write_contents(self, url, login_need):
        if login_need is False:
            try:
                captcha_img = self.driver.find_element(By.CSS_SELECTOR, "#captcha_img")
                captcha_url = captcha_img.get_attribute("src")
                cookies = self.driver.get_cookies()

                session = requests.Session()
                for cookie in cookies:
                    session.cookies.set(cookie['name'], cookie['value'])

                response = session.get(captcha_url, headers={
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
                })

                if response.status_code == 200:
                    with open("captcha.jpg", "wb") as file:
                        file.write(response.content)
                    try:
                        result = solver.normal('captcha.jpg')
                    except Exception as e:
                        sys.exit(e)
                    else:
                        self.api_count += 1
                        self.api_count_updated.emit(self.api_count)
                        write_api_count(self.api_count)
                        number = result['code']
                        self.log_updated.emit(f'[{url}] 캡챠 인식 완료 : {number}')

                        captcha_box = self.driver.find_element(By.CSS_SELECTOR, "#captcha_key")
                        captcha_box.send_keys(number)
            except NoSuchElementException:
                print("Captcha image not found. Skipping captcha solving.")

        try:
            wait = WebDriverWait(self.driver, 3)
            submit_button = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#btn_submit")))
            self.driver.execute_script("arguments[0].scrollIntoView();", submit_button)
            submit_button.click()
        except UnexpectedAlertPresentException as e:
            self.driver.quit()
            self.log_updated.emit(f'[{url}] 경고창 감지')
            self.init_web_driver()
        except Exception as e:
            print_with_debug(e)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.filename = 'urls.json'
        self.login_manage_button = QPushButton("로그인주소 관리")
        self.write_button = QPushButton("글쓰기 시작")
        self.api_count = read_api_count()
        self.title_modify_button = QPushButton("변경")
        self.add_button = QPushButton("추가")
        self.delete_button = QPushButton("삭제")
        self.table = QTableWidget()
        self.progress_bar = QProgressBar(self)
        self.log_edit_box = QTextEdit(self)
        self.count_label = None
        self.url_manager = None
        self.worker = None
        self.init_ui()
        self.load_settings()
        self.loadFilePaths()
        self.load_urls_from_file(self.filename)

    def init_ui(self):
        try:
            self.setWindowTitle('SmartWriter')
            self.setWindowIcon(QIcon('main_icon.png'))

            writing_delay_label = QLabel('글쓰기 딜레이(초)')
            self.writing_delay_input = QLineEdit()
            overall_delay_label = QLabel('전체 딜레이(초)')
            self.overall_delay_input = QLineEdit()
            repeat_label = QLabel('전체 반복 회수')
            self.repeat_input = QLineEdit()
            self.convert_checkbox = QCheckBox('특수문자 치환')

            # Set the size policy for the new input boxes
            self.writing_delay_input.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            self.overall_delay_input.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            self.repeat_input.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

            # Create a layout for the new input boxes and labels
            delay_layout = QHBoxLayout()
            delay_layout.addWidget(writing_delay_label)
            delay_layout.addWidget(self.writing_delay_input)
            delay_layout.addWidget(overall_delay_label)
            delay_layout.addWidget(self.overall_delay_input)
            delay_layout.addWidget(repeat_label)
            delay_layout.addWidget(self.repeat_input)
            delay_layout.addWidget(self.convert_checkbox)


            self.add_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
            self.delete_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

            self.url_edit_box = QLineEdit()
            self.id_edit_box = QLineEdit()
            self.pw_edit_box = QLineEdit()

            inner_layout = QVBoxLayout()
            inner_layout_top_layout = QHBoxLayout()
            inner_layout_top_layout.addWidget(QLabel('URL:'))
            inner_layout_top_layout.addWidget(self.url_edit_box, stretch=10)
            inner_layout_top_layout.addWidget(QLabel('ID:'))
            inner_layout_top_layout.addWidget(self.id_edit_box, stretch=2)
            inner_layout_top_layout.addWidget(QLabel('PW:'))
            inner_layout_top_layout.addWidget(self.pw_edit_box, stretch=2)
            inner_layout_top_layout.addWidget(self.add_button, stretch=1)
            inner_layout_top_layout.addWidget(self.delete_button, stretch=1)

            inner_layout.addLayout(inner_layout_top_layout)
            self.url_table_widget = QTableWidget()
            self.url_table_widget.setColumnCount(3)
            self.url_table_widget.setHorizontalHeaderLabels(['URL', 'ID', 'PW'])
            self.url_table_widget.setSelectionBehavior(QTableWidget.SelectRows)
            self.url_table_widget.setSelectionMode(QTableWidget.ExtendedSelection)
            self.url_table_widget.setColumnWidth(0, 800)
            self.url_table_widget.setColumnWidth(1, 100)
            self.url_table_widget.verticalHeader().setVisible(False)

            # 마지막에 이 코드를 추가하여 URL 열이 나머지 너비를 차지하도록 설정합니다.
            self.url_table_widget.horizontalHeader().setStretchLastSection(True)
            inner_layout.addWidget(self.url_table_widget)

            group_box = QGroupBox()
            group_box.setStyleSheet("QGroupBox { border: 1px solid gray; }")
            group_box.setLayout(inner_layout)

            self.add_button.clicked.connect(self.on_add_button_click)
            self.delete_button.clicked.connect(self.on_delete_button_click)

            main_layout = QVBoxLayout()
            main_layout.addLayout(delay_layout)

            button_layout = QHBoxLayout()
            button_layout.addWidget(self.login_manage_button)
            button_layout.addWidget(self.write_button)

            main_layout.addLayout(button_layout)

            self.count_label = QLabel(f"누적 캡챠 호출 횟수 : {self.api_count}", self)
            self.count_label.setAlignment(Qt.AlignCenter)
            main_layout.addWidget(self.count_label)

            self.login_manage_button.clicked.connect(self.on_login_manage_button_click)
            self.write_button.clicked.connect(self.on_write_button_click)

            load_excel_button = QPushButton('엑셀 불러오기')
            load_excel_button.clicked.connect(self.on_load_excel_button)

            self.listBox = QListWidget(self)
            self.listBox.setFixedHeight(100)

            main_layout.addWidget(group_box)
            main_layout.addWidget(self.progress_bar)
            main_layout.addWidget(load_excel_button)
            main_layout.addWidget(self.listBox)

            self.log_edit_box.setReadOnly(True)  # 읽기 전용으로 설정
            self.log_edit_box.setFixedHeight(150)  # 높이 설정
            main_layout.addWidget(self.log_edit_box)

            container = QWidget()
            container.setLayout(main_layout)
            self.setCentralWidget(container)
            self.resize(1200, 800)
            x, y = get_center_position(self.width(), self.height())
            self.move(x, y)
            self.add_log('프로그램 시작')
        except Exception as e:
            print_with_debug(e)

    def getFileList(self):
        try:
            return [self.listBox.item(i).text() for i in range(self.listBox.count())]
        except Exception as e:
            print_with_debug(e)
            return 0

    def save_settings(self):
        settings = {
            'writing_delay': self.writing_delay_input.text(),
            'overall_delay': self.overall_delay_input.text(),
            'repeat': self.repeat_input.text()
        }
        with open('settings.json', 'w') as f:
            json.dump(settings, f)

    def load_settings(self):
        try:
            with open('settings.json', 'r') as f:
                settings = json.load(f)
                self.writing_delay_input.setText(settings.get('writing_delay', ''))
                self.overall_delay_input.setText(settings.get('overall_delay', ''))
                self.repeat_input.setText(settings.get('repeat', ''))
        except Exception as e:
            print_with_debug(e)

    def add_log(self, log):
        current_time = datetime.datetime.now()
        formatted_time = current_time.strftime("[%Y-%m-%d %H:%M:%S]")
        formatted_log = f"{formatted_time} {log}"
        print(formatted_log)
        self.log_edit_box.append(formatted_log)
        self.log_edit_box.verticalScrollBar().setValue(self.log_edit_box.verticalScrollBar().maximum())

    def addFilePaths(self, filePaths):
        try:
            for filePath in filePaths:
                if filePath not in [self.listBox.item(i).text() for i in range(self.listBox.count())]:
                    self.listBox.addItem(filePath)
            self.saveFilePaths()
        except Exception as e:
            print_with_debug(e)

    def saveFilePaths(self):
        try:
            filePaths = [self.listBox.item(i).text() for i in range(self.listBox.count())]
            with open('file_paths.json', 'w') as file:
                json.dump(filePaths, file)
        except Exception as e:
            print_with_debug(e)

    def loadFilePaths(self):
        try:
            with open('file_paths.json', 'r') as file:
                filePaths = json.load(file)
                self.addFilePaths(filePaths)
        except (FileNotFoundError, json.JSONDecodeError):
            pass

    def on_load_excel_button(self):
        try:
            options = QFileDialog.Options()
            options |= QFileDialog.DontUseNativeDialog
            files, _ = QFileDialog.getOpenFileNames(self, "Select Excel Files", "",
                                                    "Excel Files (*.xlsx *.xls);;All Files (*)", options=options)

            self.listBox.clear()
            if files:
                self.addFilePaths(files)
        except Exception as e:
            print_with_debug(e)

    def on_login_manage_button_click(self):
        try:
            self.url_manager = URLManager()
            self.url_manager.show()
        except Exception as e:
            print_with_debug(e)

    def on_write_button_click(self):
        self.save_settings()
        entries = self.get_all_entries()
        self.progress_bar.setMaximum(len(entries))
        self.worker = Worker(entries, int(self.writing_delay_input.text()), int(self.overall_delay_input.text()),
                             int(self.repeat_input.text()), self.convert_checkbox.isChecked(), self.getFileList())
        self.worker.progress_updated.connect(self.update_progress)
        self.worker.log_updated.connect(self.add_log)
        self.worker.api_count_updated.connect(self.update_api_count)
        self.worker.start()

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def update_api_count(self, count):
        self.api_count = count
        self.count_label.setText(f"누적 캡챠 호출 횟수 : {self.api_count}")

    def on_add_button_click(self):
        try:
            url = self.url_edit_box.text()
            user_id = self.id_edit_box.text()
            password = self.pw_edit_box.text()

            if url:  # URL 필수 입력
                row_position = self.url_table_widget.rowCount()
                self.url_table_widget.insertRow(row_position)
                url_item = QTableWidgetItem(url)
                url_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)  # Read-only 설정
                self.url_table_widget.setItem(row_position, 0, url_item)

                if user_id:
                    user_id_item = QTableWidgetItem(user_id)
                    user_id_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)  # Read-only 설정
                    self.url_table_widget.setItem(row_position, 1, user_id_item)

                if password:
                    password_item = QTableWidgetItem(password)
                    password_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)  # Read-only 설정
                    self.url_table_widget.setItem(row_position, 2, password_item)

                # 입력란 초기화
                self.url_edit_box.clear()
                self.id_edit_box.clear()
                self.pw_edit_box.clear()

                self.write_urls_to_file(self.filename)
        except Exception as e:
            print_with_debug(e)

    def load_urls_from_file(self, filename):
        try:
            with open(filename, 'r') as file:
                data = json.load(file)
                urls = data.get('urls', [])
                for url_info in urls:
                    row_position = self.url_table_widget.rowCount()
                    self.url_table_widget.insertRow(row_position)

                    url_item = QTableWidgetItem(url_info.get('url', ''))
                    url_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)  # Read-only 설정
                    self.url_table_widget.setItem(row_position, 0, url_item)

                    id_item = QTableWidgetItem(url_info.get('id', ''))
                    id_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)  # Read-only 설정
                    self.url_table_widget.setItem(row_position, 1, id_item)

                    pw_item = QTableWidgetItem(url_info.get('pw', ''))
                    pw_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)  # Read-only 설정
                    self.url_table_widget.setItem(row_position, 2, pw_item)
        except FileNotFoundError:
            pass  # 파일이 없는 경우 무시
        except Exception as e:
            print_with_debug(e)

    def write_urls_to_file(self, filename):
        try:
            urls = []
            for row in range(self.url_table_widget.rowCount()):
                url = self.url_table_widget.item(row, 0).text() if self.url_table_widget.item(row, 0) else ''
                user_id = self.url_table_widget.item(row, 1).text() if self.url_table_widget.item(row, 1) else ''
                password = self.url_table_widget.item(row, 2).text() if self.url_table_widget.item(row, 2) else ''
                urls.append({'url': url, 'id': user_id, 'pw': password})

            data = {'urls': urls}
            with open(filename, 'w') as file:
                json.dump(data, file)
        except Exception as e:
            print_with_debug(e)

    def on_delete_button_click(self):
        try:
            selected_rows = self.url_table_widget.selectionModel().selectedRows()
            for row in sorted(selected_rows, reverse=True):
                self.url_table_widget.removeRow(row.row())

            self.write_urls_to_file(self.filename)
        except Exception as e:
            print_with_debug(e)
    def get_all_entries(self):
        try:
            entries = []
            for row in range(self.url_table_widget.rowCount()):
                url = self.url_table_widget.item(row, 0).text() if self.url_table_widget.item(row, 0) else ''
                user_id = self.url_table_widget.item(row, 1).text() if self.url_table_widget.item(row, 1) else ''
                password = self.url_table_widget.item(row, 2).text() if self.url_table_widget.item(row, 2) else ''
                entries.append([url, user_id, password])
            return entries
        except Exception as e:
            print_with_debug(e)


class URLManager(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('URL Manager')
        self.setGeometry(100, 100, 800, 600)
        self.centerWindow()

        self.layout = QVBoxLayout()

        self.input_layout = QHBoxLayout()
        self.write_url_edit = QLineEdit(self)
        self.write_url_edit.setPlaceholderText('write_url')
        self.login_url_edit = QLineEdit(self)
        self.login_url_edit.setPlaceholderText('login_url')
        self.add_button = QPushButton('추가', self)
        self.delete_button = QPushButton('삭제', self)
        self.input_layout.addWidget(self.write_url_edit)
        self.input_layout.addWidget(self.login_url_edit)
        self.input_layout.addWidget(self.add_button)
        self.input_layout.addWidget(self.delete_button)

        self.layout.addLayout(self.input_layout)

        self.table_widget = QTableWidget(self)
        self.table_widget.setColumnCount(2)
        self.table_widget.setColumnWidth(0, 400)
        self.table_widget.setHorizontalHeaderLabels(['write_url', 'login_url'])

        self.table_widget.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table_widget.verticalHeader().setVisible(False)  # 인덱스 번호 숨기기
        self.layout.addWidget(self.table_widget)

        self.setLayout(self.layout)

        self.add_button.clicked.connect(self.add_url)
        self.delete_button.clicked.connect(self.delete_url)

        self.load_data()

    def centerWindow(self):
        qr = self.frameGeometry()
        cp = QApplication.desktop().screen().rect().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def add_url(self):
        write_url = self.write_url_edit.text()
        login_url = self.login_url_edit.text()

        if write_url and login_url:
            row_position = self.table_widget.rowCount()
            self.table_widget.insertRow(row_position)
            self.table_widget.setItem(row_position, 0, QTableWidgetItem(write_url))
            self.table_widget.setItem(row_position, 1, QTableWidgetItem(login_url))

            self.save_data()

            self.write_url_edit.clear()
            self.login_url_edit.clear()
        else:
            QMessageBox.warning(self, 'Error', 'Please enter both write_url and login_url')

    def delete_url(self):
        selected_row = self.table_widget.currentRow()
        if selected_row >= 0:
            self.table_widget.removeRow(selected_row)
            self.save_data()
        else:
            QMessageBox.warning(self, 'Error', 'Please select a row to delete')

    def save_data(self):
        data = []
        for row in range(self.table_widget.rowCount()):
            write_url = self.table_widget.item(row, 0).text()
            login_url = self.table_widget.item(row, 1).text()
            data.append({'write_url': write_url, 'login_url': login_url})

        with open('login_urls.json', 'w') as f:
            json.dump(data, f)

    def load_data(self):
        try:
            with open('login_urls.json', 'r') as f:
                data = json.load(f)
                for entry in data:
                    row_position = self.table_widget.rowCount()
                    self.table_widget.insertRow(row_position)
                    self.table_widget.setItem(row_position, 0, QTableWidgetItem(entry['write_url']))
                    self.table_widget.setItem(row_position, 1, QTableWidgetItem(entry['login_url']))
        except FileNotFoundError:
            with open('login_urls.json', 'w') as f:
                json.dump([], f)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    qtmodern.styles.dark(app)
    main_win = MainWindow()
    main_win.show()
    sys.exit(app.exec_())