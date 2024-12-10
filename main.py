import configparser
import sys
import os
import datetime
import json
import re
import urllib.parse
import zipfile
import itertools
import shutil
import gc
import threading
import numpy as np
import capsolver
import httpx
import subprocess
import psutil
import pyautogui


from fake_useragent import UserAgent
from urllib.parse import urlparse, parse_qs
from concurrent.futures import ThreadPoolExecutor, as_completed
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities

from PyQt5 import QtWidgets


from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QLineEdit, QMenuBar, \
    QAction, QMessageBox, QMainWindow, QMenu, QProgressBar, QTextEdit, QActionGroup, QSizePolicy, QGroupBox, \
    QTableWidget, QComboBox, \
    QGridLayout, QTableWidget, QHeaderView, QAbstractItemView, QTableWidgetItem, QCheckBox, QListWidget, QFileDialog

from PyQt5.QtGui import QIcon
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options

import traceback
import qtmodern.styles
import qtmodern.windows
from PyQt5 import QtWidgets as qtw
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from selenium.common.exceptions import TimeoutException

#from selenium import webdriver
from seleniumwire import webdriver  # Selenium Wire 사용
from selenium.webdriver.common.by import By
from selenium.common.exceptions import UnexpectedAlertPresentException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
import requests
import time
from twocaptcha.solver import TwoCaptcha
import pandas as pd
import pyperclip
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.proxy import Proxy, ProxyType

import random

# 한글 초성, 중성, 종성 리스트
chosung = ['ㄱ', 'ㄲ', 'ㄴ', 'ㄷ', 'ㄸ', 'ㄹ', 'ㅁ', 'ㅂ', 'ㅃ', 'ㅅ', 'ㅆ', 'ㅇ', 'ㅈ', 'ㅉ', 'ㅊ', 'ㅋ', 'ㅌ', 'ㅍ', 'ㅎ']
jungsung = ['ㅏ', 'ㅐ', 'ㅑ', 'ㅒ', 'ㅓ', 'ㅔ', 'ㅕ', 'ㅖ', 'ㅗ', 'ㅘ', 'ㅙ', 'ㅚ', 'ㅛ', 'ㅜ', 'ㅝ', 'ㅞ', 'ㅟ', 'ㅠ', 'ㅡ', 'ㅢ', 'ㅣ']
jongsung = [''] + ['ㄱ', 'ㄲ', 'ㄳ', 'ㄴ', 'ㄵ', 'ㄶ', 'ㄷ', 'ㄹ', 'ㄺ', 'ㄻ', 'ㄼ', 'ㄽ', 'ㄾ', 'ㄿ', 'ㅀ', 'ㅁ', 'ㅂ', 'ㅄ', 'ㅅ', 'ㅆ',
                   'ㅇ', 'ㅈ', 'ㅊ', 'ㅋ', 'ㅌ', 'ㅍ', 'ㅎ']


config = configparser.ConfigParser()
config.read('config.ini')

# GitHub 토큰 불러오기
GITHUB_TOKEN = config['github']['token']

# IP 확인을 위한 URL (외부에서 접속한 IP를 반환해줌)
ip_check_url = "https://httpbin.org/ip"

# 접속 테스트할 일반 웹사이트 URL
test_site = "https://www.google.com"

chrome_config_path = "chrome_config.json"
with open(chrome_config_path, "r") as file:
    config = json.load(file)

# 크롬 실행 경로 가져오기
chrome_path = config.get("chrome_path")
if not chrome_path:
    raise FileNotFoundError("Chrome 실행 경로가 JSON 파일에 지정되지 않았습니다.")

def get_base_domain(url):
    parsed_url = urlparse(url)
    domain = parsed_url.netloc
    if domain.startswith('www.'):
        domain = domain[4:]

    base_domain = domain.split('/')[0]
    return base_domain

def get_full_base_domain(url):
    base_domain = get_base_domain(url)
    return f'https://{base_domain}'

def get_sub_path(url):
    parsed_url = urlparse(url)
    domain = parsed_url.netloc
    path = parsed_url.path
    if path.startswith('/'):
        sub_path = path.split('/')[1]  # 첫 번째 '/' 이후의 부분을 추출
    else:
        sub_path = ''

    return sub_path

def get_board_value(url):
    parsed_url = urlparse(url)
    query_params = parse_qs(parsed_url.query)
    if 'board' in query_params:
        return query_params['board'][0]
    else:
        return None

def download_and_extract_zip(url, extract_to):
    headers = {
        'Authorization': f'token {GITHUB_TOKEN}',
        'Accept': 'application/octet-stream'
    }
    session = requests.Session()
    session.headers.update(headers)

    response = session.get(url, stream=True)

    if response.status_code != 200:
        print(f"Failed to download file. Status code: {response.status_code}")
        return False

    zip_path = os.path.join(extract_to, "update.zip")
    with open(zip_path, "wb") as file:
        for chunk in response.iter_content(chunk_size=8192):
            file.write(chunk)

    try:
        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            zip_ref.extractall(extract_to)
    except zipfile.BadZipFile:
        print("Downloaded file is not a valid zip file")
        return False

    os.remove(zip_path)
    return True


def get_latest_release():
    repo_owner = 'YunJunYoung'
    repo_name = 'smartwriter'
    url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/releases/latest"
    headers = {
        'Authorization': f'token {GITHUB_TOKEN}',
        'Accept': 'application/vnd.github.v3+json'
    }
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        release_info = response.json()
        latest_version = release_info["tag_name"]
        download_url = release_info["assets"][0]["browser_download_url"]
        return latest_version, download_url
    else:
        print(f"Failed to fetch release information. Status code: {response.status_code}")
        print(f"Response content: {response.content}")

def get_domain(url):
    parsed_url = urlparse(url)
    domain = f"{parsed_url.scheme}://{parsed_url.netloc}"
    return domain

def extract_base_url(url: str) -> str:
    # 쿼리 파라미터 구분자 위치 찾기
    query_pos = url.find('?q')

    # '?q'가 없다면 전체 URL을 반환
    if query_pos == -1:
        return url

    # '?q' 직전까지의 URL 반환
    return url[:query_pos]


def compare_versions(v1, v2):
    """
    Compare two versions. Return True if v2 is newer than v1.
    Versions should be in the format 'v1.2.3' or '1.2.3'.
    """

    def normalize(version):
        return [float(x) for x in version.strip('v').split('.')]

    return normalize(v1) < normalize(v2)


def load_version():
    with open('version.json', 'r') as file:
        data = json.load(file)
        return data.get('version', 'Unknown version')

def read_proxy_list(file_path):
    with open(file_path, 'r') as file:
        proxies = file.readlines()
    proxy_list = [proxy.strip() for proxy in proxies]
    return proxy_list

def read_proxy_user(file_path):
    with open(file_path, 'r') as file:
        user_data = json.load(file)
    return user_data['username'], user_data['password']

# Usage example
PROXY_LIST = read_proxy_list('proxy_list.txt')
PROXY_USERNAME, PROXY_PASSWORD = read_proxy_user('proxy_user.json')
proxy_iterator = itertools.cycle(PROXY_LIST)

# Function to get the next proxy IP
def get_proxy_ip():
   return next(proxy_iterator)

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


def generate_random_number(length):
    # 랜덤한 숫자를 생성하고 그 숫자의 길이를 조정합니다.
    number = random.randint(10 ** (length - 1), 10 ** length - 1)
    return number

custom_file_name = 'custom_writer.txt'
def generate_random_writer():
    with open(custom_file_name, 'r', encoding='utf-8') as file:
        words = file.read().splitlines()  # 파일을 줄 단위로 읽어와서 리스트로 변환

    if words:
        return random.choice(words)  # 리스트에서 랜덤으로 한 단어 선택
    else:
        return None  # 파일이 비어있을 경우 None 반환

def generate_random_chinese_characters(length):
    if length < 4:
        length = 4
    elif length > 6:
        length = 6

    # 중국어 글자의 유니코드 범위 설정 (기본 범위 CJK Unified Ideographs)
    start, end = 0x4E00, 0x9FFF

    # 랜덤한 중국어 글자 생성
    chinese_characters = ''.join(chr(random.randint(start, end)) for _ in range(length))
    return chinese_characters


def generate_random_japanese_characters(length):
    if length < 4:
        length = 4
    elif length > 6:
        length = 6

    # 히라가나와 가타카나의 유니코드 범위 설정
    hiragana_start, hiragana_end = 0x3040, 0x309F
    katakana_start, katakana_end = 0x30A0, 0x30FF

    # 랜덤한 일본어 글자 생성
    japanese_characters = ''.join(chr(random.choice(
        list(range(hiragana_start, hiragana_end + 1)) +
        list(range(katakana_start, katakana_end + 1))
    )) for _ in range(length))

    return japanese_characters


def generate_english_name():
    consonants = "bcdfghjklmnpqrstvwxyz"
    vowels = "aeiou"

    # 이름 길이는 6~8 글자로 랜덤하게 선택
    name_length = random.choice(range(6, 9))
    name = ''

    # 자음과 모음을 번갈아가면서 생성
    for i in range(name_length):
        if i % 2 == 0:  # 짝수 위치에는 자음
            name += random.choice(consonants)
        else:  # 홀수 위치에는 모음
            name += random.choice(vowels)

    # 첫 글자는 대문자로 변환
    return name.capitalize()


def is_file_name(input_string):
    # 파일 이름인 경우 마지막이 .txt로 끝나는지 체크
    if input_string.endswith('.txt'):
        return True
    else:
        return False


def read_whole_text(file_path):
    try:
        # 현재 스크립트의 경로를 가져옵니다.
        current_dir = os.path.dirname(__file__)

        # txt 폴더 아래의 파일 경로를 만듭니다.
        full_path = os.path.join(current_dir, 'txt', file_path)

        with open(full_path, 'r', encoding='utf-8') as file:
            content = file.read()
            return content
    except FileNotFoundError:
        return f"File not found: {full_path}"
    except Exception as e:
        return f"Error reading file: {str(e)}"

def text_to_html(text):
    try:
        #html_content = "<html>\n<body>\n"
        html_content = text.replace('\n', '<br>\n')
        #html_content += "\n</body>\n</html>"
        return html_content
    except Exception as e:
        print_with_debug(e)


def print_with_debug(msg):
    print(msg)
    traceback.print_exc()


def get_center_position(width, height):
    active_screen = QApplication.desktop().screenNumber(QApplication.desktop().cursor().pos())
    screen_geometry = QApplication.desktop().screenGeometry(active_screen)
    x = int((screen_geometry.width() - width) / 2 + screen_geometry.left())
    y = int((screen_geometry.height() - height) / 2 + screen_geometry.top())
    return x, y


class ExcelRandomPicker:
    def __init__(self, excel_file_list):
        self.excel_file_list = excel_file_list
        self.data_frames = self.read_all_excel_files()

    def read_all_excel_files(self):
        data_frames = []
        for file in self.excel_file_list:
            # 첫 번째 행을 건너뛰지 않고 파일을 읽음
            df = pd.read_excel(file)
            # 열 이름을 수동으로 설정
            df.columns = ['No', '제목', '내용', '이미지URL', '사이트URL', '이름']
            data_frames.append(df)
        return data_frames

    def get_random_titles_contents(self, N):
        total_files = len(self.excel_file_list)
        if N % total_files != 0:
            raise ValueError("N must be divisible by the number of files")

        per_file_count = N // total_files
        all_random_titles_contents = []

        for i in range(per_file_count):
            for df in self.data_frames:
                # 무작위로 행 선택
                random_index = random.randint(0, len(df) - 1)
                row = df.iloc[random_index]

                title = row.get('제목')
                content = row.get('내용')
                img_url = row.get('이미지URL')
                site_url = row.get('사이트URL')
                name = row.get('이름')

                if pd.isna(title) or pd.isna(content):
                    continue

                # img_url이 nan인지 확인하여 None으로 변경
                if pd.isna(img_url):
                    img_url = None

                if pd.isna(site_url):
                    site_url = None

                all_random_titles_contents.append((title, content, img_url, site_url, name))

                # 디버깅: 현재 리스트 길이 출력

        # 총 N개가 반환되었는지 확인
        if len(all_random_titles_contents) != N:
            print(f"Warning: Only {len(all_random_titles_contents)} items were collected.")

        return all_random_titles_contents


class Worker(QThread):
    progress_updated = pyqtSignal(int)
    log_updated = pyqtSignal(str)
    api_count_updated = pyqtSignal(int)

    def __init__(self, entries, writing_delay, overall_delay, repeat, convert, excel_file_list, name_language,
                 num_tabs, use_proxy, use_chrome, x_pos, y_pos):
        super().__init__()
        try:
            self.entries = entries
            self.writing_delay = writing_delay
            self.overall_delay = overall_delay
            self.repeat = repeat
            self.convert = convert
            self.excel_file_list = excel_file_list
            self.picker = ExcelRandomPicker(excel_file_list)
            self.current_file_index = 0
            self.name_language = name_language
            self.num_tabs = num_tabs
            self.total_number = len(self.entries) * repeat * num_tabs
            random_titles_contents = self.picker.get_random_titles_contents(self.total_number)
            self.titles, self.contents, self.img_urls, self.site_urls, self.names = zip(*random_titles_contents)
            self.write_index = 0
            self.use_proxy = use_proxy
            self.temp_profiles = []
            self.driver = None
            self.session = None
            self.use_chrome = use_chrome
            self.x_pos = x_pos
            self.y_pos = y_pos
        except Exception as e:
            print_with_debug(e)

    def get_name(self, name_language):
        if name_language == '한글':
            name = generate_korean_name()
        elif name_language == '영어':
            name = generate_english_name()
        elif name_language == '한글+숫자':
            name = f'{generate_korean_name()}{generate_random_number(4)}'
        elif name_language == '영어+숫자':
            name = f'{generate_english_name()}{generate_random_number(3)}'
        elif name_language == '숫자만':
            name = generate_random_number(6)
        elif name_language == '중국어':
            name = generate_random_chinese_characters(4)
        elif name_language == '일본어':
            name = generate_random_japanese_characters(6)
        elif name_language == '커스텀':
            name = generate_random_writer()
        else:
            name = name_language

        return name

    def perform_task(self, repeat, index, sub_index, url):
        self.log_updated.emit(f'\n[Index:{repeat + 1}_{index + 1}_{sub_index + 1}] [{url}] 글쓰기 시작')
        try:
            title = self.titles[self.write_index]
            content = self.contents[self.write_index]
            img_url = self.img_urls[self.write_index]
            site_url = self.site_urls[self.write_index]
            name = self.names[self.write_index]

            title = str(title)
            name = str(name)
            content = str(content)
            name = self.get_name(name[1:])

            if is_file_name(content):
                text_content = read_whole_text(content)
                content = text_to_html(text_content)

            if img_url:
                if site_url:
                    edit_content = f"<a href=\"{site_url}\" target=\"_blank\"><img src=\"{img_url}\"><p>{content}</p></a>"
                else:
                    edit_content = f"<img src=\"{img_url}\"><p>{content}</p>"
            else:
                edit_content = f"<p>{content}</p>"

            if self.use_chrome:
                ret = self.write_contents_for_cloud_flare(url, name, title, edit_content, repeat, index, sub_index)
            else:
                ret = self.write_contents(url, name, title, edit_content, repeat, index, sub_index)
            self.write_index += 1

            if ret:
                self.log_updated.emit(f'[Index:{repeat + 1}_{index + 1}_{sub_index + 1}] [{url}] 글쓰기 성공')
            else:
                self.log_updated.emit(f'[Index:{repeat + 1}_{index + 1}_{sub_index + 1}] [{url}] 글쓰기 실패')
            return ret
        except Exception as e:
            print_with_debug(e)
            self.log_updated.emit(f'[Index:{repeat + 1}_{index + 1}_{sub_index+1}] [{url}] 글쓰기 실패')
            return False

    def save_cookies_for_cloud_flare(self, url, repeat, index, sub_index):
        try:
            full_base_url = get_full_base_domain(url)
            # 프록시 설정을 초기화 시점에 설정
            proxies = None
            if self.use_proxy:
                proxy_ip = get_proxy_ip()
                proxy_host, proxy_port = proxy_ip.split(":")[0], proxy_ip.split(":")[1]
                proxy = f'http://{PROXY_USERNAME}:{PROXY_PASSWORD}@{proxy_host}:{proxy_port}'
                proxies = {
                    'http://': proxy,
                    'https://': proxy
                }

            # HTTP/2 지원을 위해 httpx 클라이언트 초기화 시 프록시 설정 포함
            self.client = httpx.Client(http2=True, proxies=proxies)
            self.log_updated.emit(f"[Index:{repeat + 1}_{index + 1}_{sub_index + 1}] [STEP] [0] [세션 정보 획득 성공]")
        except Exception as e:
            self.log_updated.emit(f"[Index:{repeat + 1}_{index + 1}_{sub_index + 1}] [STEP] [0] 에러발생 : {e}")

    def start_browser_with_subprocess(self, url, user_data_dir, debugging_port=9222):
       try:
           """Subprocess로 Chrome 실행 및 인증"""
           subprocess.Popen([
               chrome_path,
               f"--remote-debugging-port={debugging_port}",
               "--no-first-run",
               "--no-default-browser-check",
               "--lang=ko-KR",  # 한국어 설정
               "--start-maximized",
               url
           ])

           # 브라우저 인증을 위한 시간 대기
           time.sleep(5)  # 수동 인증을 완료할 시간
           pyautogui.click(self.x_pos, self.y_pos, duration=0.5)
           time.sleep(1)  # 수동 인증을 완료할 시간

       except Exception as e:
           self.log_updated.emit(e)

    def connect_to_existing_browser(self, debugging_port=9222):
        """Selenium으로 인증된 브라우저 세션에 연결"""
        try:
            chrome_options = Options()
            chrome_options.add_experimental_option("debuggerAddress", f"127.0.0.1:{debugging_port}")
            self.driver = webdriver.Chrome(options=chrome_options)
            self.log_updated.emit('크롬실행. 연결 완료')
        except Exception as e:
            self.log_updated.emit(e)

    def _set_value_with_javascript(self, driver, element, text):
        try:
            # JavaScript를 사용하여 값을 설정하고 입력 이벤트를 트리거합니다.
            driver.execute_script("""
            arguments[0].value = arguments[1];
            arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            """, element, text)
            time.sleep(0.1)
        except Exception as e:
            print(f"Error: {e}")

    def save_cookies(self, url, repeat, index, sub_index):
        try:
            full_base_url = get_full_base_domain(url)
            # 프록시 설정을 초기화 시점에 설정
            proxies = None
            if self.use_proxy:
                proxy_ip = get_proxy_ip()
                proxy_host, proxy_port = proxy_ip.split(":")[0], proxy_ip.split(":")[1]
                proxy = f'http://{PROXY_USERNAME}:{PROXY_PASSWORD}@{proxy_host}:{proxy_port}'
                proxies = {
                    'http://': proxy,
                    'https://': proxy
                }

            # HTTP/2 지원을 위해 httpx 클라이언트 초기화 시 프록시 설정 포함
            self.client = httpx.Client(http2=True, proxies=proxies)

            # Selenium에서 쿠키 가져오기
            cookies = self.driver.get_cookies()

            # 각 쿠키에 domain과 path를 설정하여 httpx 클라이언트에 추가
            for cookie in cookies:
                self.client.cookies.set(
                    name=cookie['name'],
                    value=cookie['value'],
                    domain=full_base_url,  # 방문 중인 도메인
                    path=cookie.get('path', '/')
                )

            self.log_updated.emit(f"[Index:{repeat + 1}_{index + 1}_{sub_index + 1}] [STEP] [0] [세션 정보 획득 성공]")
        except Exception as e:
            self.log_updated.emit(f"[Index:{repeat + 1}_{index + 1}_{sub_index + 1}] [STEP] [0] 에러발생 : {e}")

    def check_post_client_token(self, url, repeat, index, sub_index):
        try:
            host = get_base_domain(url)
            full_base_url = get_full_base_domain(url)
            post_client_token_url = f'{full_base_url}/set_post_client_token.cm'

            # 브라우저와 일치하도록 설정한 헤더
            post_client_token_headers = {
                "accept": "application/json, text/javascript, */*; q=0.01",
                "accept-encoding": "gzip, deflate, br, zstd",
                "accept-language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
                "origin": full_base_url,
                "referer": url,
                'sec-ch-ua': '"Whale";v="3", "Not-A.Brand";v="8", "Chromium";v="126"',
                "sec-ch-ua-mobile": "?0",
                'sec-ch-ua-platform': '"Windows"',
                "sec-fetch-dest": "empty",
                "sec-fetch-mode": "cors",
                "sec-fetch-site": "same-origin",
                "x-requested-with": "XMLHttpRequest",
                "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Whale/3.28.266.14 Safari/537.36"
            }

            # 클라이언트를 통해 HTTP/2 요청 수행
            response = self.client.post(post_client_token_url, headers=post_client_token_headers, timeout=10)

            # 응답 내용 처리
            escaped_text = response.text.replace("<", "&lt;").replace(">", "&gt;")
            if response.status_code == 200:
                self.log_updated.emit(
                    f'[Index:{repeat + 1}_{index + 1}_{sub_index + 1}] [STEP] [2] [유효성 검사] [응답 : {response.status_code}] [내용 : {escaped_text}]')
                return True
            else:
                self.log_updated.emit(
                    f'[Index:{repeat + 1}_{index + 1}_{sub_index + 1}] [STEP] [2] [유효성 검사] [응답 : {response.status_code}] [내용 : {escaped_text}]')
                return False
        except Exception as e:
            self.log_updated.emit(f'[Index:{repeat + 1}_{index + 1}_{sub_index + 1}] [STEP] [2] 에러발생 : {e}')

    def post_add(self, write_token, write_token_key, url, content, name, title, repeat, index, sub_index):
        try:
            sub_path = get_sub_path(url)
            menu_url = f'/{sub_path}/'
            board_code = get_board_value(url)
            host = get_base_domain(url)
            full_base_url = get_full_base_domain(url)
            post_url = f'{full_base_url}/backpg/post_add.cm'

            request_data = {
                'idx': '0',
                'menu_url': menu_url,
                'back_url': '',
                'back_page_num': '',
                'board_code': board_code,
                'body': content,
                'plain_body': '',
                'is_editor': 'Y',
                'represent_img': '',
                'img': '',
                'img_tmp_no': '',
                'is_notice': 'no',
                'category_type': '1',
                'write_token': write_token,
                'write_token_key': write_token_key,
                'is_secret_post': 'no',
                'nick': name,
                'secret_pass': '14143651',
                'subject': title
            }

            # 헤더 설정
            post_client_token_headers = {
                "accept": "application/json, text/javascript, */*; q=0.01",
                "accept-encoding": "gzip, deflate, br, zstd",
                "accept-language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
                "origin": full_base_url,
                "referer": url,
                'sec-ch-ua': '"Whale";v="3", "Not-A.Brand";v="8", "Chromium";v="126"',
                "sec-ch-ua-mobile": "?0",
                'sec-ch-ua-platform': '"Windows"',
                "sec-fetch-dest": "empty",
                "sec-fetch-mode": "cors",
                "sec-fetch-site": "same-origin",
                "x-requested-with": "XMLHttpRequest",
                "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Whale/3.28.266.14 Safari/537.36",
                "cache-control": "max-age=0",
                "content-type": "application/x-www-form-urlencoded",
                "upgrade-insecure-requests": "1"
            }

            # 클라이언트를 통해 HTTP/2 요청 수행
            response = self.client.post(post_url, data=request_data, headers=post_client_token_headers, timeout=10)

            # 응답 내용 처리
            escaped_text = response.text.replace("<", "&lt;").replace(">", "&gt;")
            if response.status_code == 200:
                self.log_updated.emit(
                    f'[Index:{repeat + 1}_{index + 1}_{sub_index + 1}] [STEP] [3] [글작성] [성공] [응답 : {response.status_code}] [내용 : {escaped_text}]')
                return True
            else:
                self.log_updated.emit(
                    f'[Index:{repeat + 1}_{index + 1}_{sub_index + 1}] [STEP] [3] [글작성] [실패] [응답 : {response.status_code}] [내용 : {escaped_text}]')
                return False
        except Exception as e:
            self.log_updated.emit(f'[Index:{repeat + 1}_{index + 1}_{sub_index + 1}] [STEP] [3] 에러발생 : {e}')
            return False

    def write_contents(self, url, name, title, content, repeat, index, sub_index):
        try:
            base_url = extract_base_url(url)

            # 첫번째만 크롬으로 url 이동
            if sub_index == 0:
                self.driver.get(url)
                # write_button_selector = "div.btn-block-right > a"
                # wait = WebDriverWait(self.driver, 10)  # 최대 10초 대기
                # write_button_element = wait.until(
                #     EC.presence_of_element_located((By.CSS_SELECTOR, write_button_selector)))
                #
                # # 스크롤을 시도하면서 클릭 가능한지 반복 확인
                # is_clicked = False
                # attempts = 0
                #
                # while not is_clicked and attempts < 5:  # 최대 5번 시도
                #     try:
                #         # 스크롤 시도
                #         self.driver.execute_script("arguments[0].scrollIntoView();", write_button_element)
                #         write_button_element = wait.until(
                #             EC.element_to_be_clickable((By.CSS_SELECTOR, write_button_selector)))
                #         self.driver.execute_script("arguments[0].click();", write_button_element)
                #         is_clicked = True  # 성공적으로 클릭되면 반복 중단
                #     except Exception as e:
                #         attempts += 1
                #         self.log_updated.emit(f"[스크롤 시도 {attempts}] 요소를 찾지 못했습니다. 오류: {str(e)}")
                #         time.sleep(1)  # 다음 시도 전 대기
                # 세션 저장 후 글쓰기 필요
                time.sleep(3)

            self.save_cookies(url, repeat, index, sub_index)
            time.sleep(1)
            write_token, write_token_key = self.get_make_token(url, repeat, index, sub_index)
            self.check_post_client_token(url, repeat, index, sub_index)
            return self.post_add(write_token, write_token_key, url, content, name, title, repeat, index, sub_index)

        except Exception as e:
            print_with_debug(e)
            return False

    def write_contents_for_cloud_flare(self, url, name, title, content, repeat, index, sub_index):
        try:
            # 새로운 브라우저 실행 및 연결
            debugging_port = 9222
            user_data_dir = r"D:\chrome_temp"
            self.start_browser_with_subprocess(url, user_data_dir, debugging_port)
            self.connect_to_existing_browser(debugging_port)
            self.save_cookies_for_cloud_flare(url, repeat, index, sub_index)
            self.check_post_client_token_for_cloud_flare(url, repeat, index, sub_index)
            ret = self.post_add_for_cloud_flare(url, content, name, title, repeat, index, sub_index)
            self.cleanup_browser()
            return ret

        except Exception as e:
            print_with_debug(e)
            return False

    def cleanup_browser(self):
        """
        브라우저 종료를 개선하여 비정상 종료 메시지를 방지
        """
        # 1. Selenium WebDriver 종료
        if hasattr(self, "driver") and self.driver:
            self.driver.quit()
            self.driver = None
        try:
            for proc in psutil.process_iter(['pid', 'name']):
                if proc.info['name'] == 'chrome.exe':
                    proc.terminate()  # 하위 프로세스도 종료
                    proc.wait(timeout=5)
        except Exception as e:
            print(f"Error during Chrome and children termination: {e}")

    def check_current_ip(self):
        try:
            # IP 주소를 확인하기 위한 테스트 요청
            response = self.client.get("https://httpbin.org/ip")

            if response.status_code == 200:
                ip_info = response.json().get("origin", "IP 주소 확인 실패")
                print(f"[STEP] 현재 요청되는 IP 주소: {ip_info}")
            else:
                print(f"[STEP] IP 주소 확인 실패: 응답 코드 {response.status_code}")

        except Exception as e:
            print(f"[STEP] IP 확인 중 오류 발생: {e}")

    def get_make_token(self, url, repeat, index, sub_index):
        try:
            #self.check_current_ip()
            host = get_base_domain(url)
            full_base_url = get_full_base_domain(url)
            make_token_url = f'{full_base_url}/ajax/make_tokens.cm'

            request_data = {
                'expire': '86400',
                'count': '1'
            }

            # 브라우저와 일치하도록 설정한 헤더
            make_tokens_headers = {
                "accept": "application/json, text/javascript, */*; q=0.01",
                "accept-encoding": "gzip, deflate, br, zstd",
                "accept-language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
                "content-type": "application/x-www-form-urlencoded; charset=UTF-8",
                "origin": full_base_url,
                "referer": url,
                'sec-ch-ua': '"Whale";v="3", "Not-A.Brand";v="8", "Chromium";v="126"',
                "sec-ch-ua-mobile": "?0",
                'sec-ch-ua-platform': '"Windows"',
                "sec-fetch-dest": "empty",
                "sec-fetch-mode": "cors",
                "sec-fetch-site": "same-origin",
                "x-requested-with": "XMLHttpRequest",
                "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Whale/3.28.266.14 Safari/537.36"
            }

            # 클라이언트 요청 수행
            response = self.client.post(make_token_url, data=request_data, headers=make_tokens_headers)
            escaped_text = response.text.replace("<", "&lt;").replace(">", "&gt;")
            if response.status_code == 200:
                tokens = response.json()
                if tokens.get('tokens'):
                    write_token = tokens['tokens'][0].get('token')
                    write_token_key = tokens['tokens'][0].get('token_key')
                    self.log_updated.emit(
                        f'[Index:{repeat + 1}_{index + 1}_{sub_index + 1}] [STEP] [1] [글작성 토큰] [성공] [token : {write_token}] [key : {write_token_key}]')
                    return write_token, write_token_key
            else:
                self.log_updated.emit(f'[Index:{repeat + 1}_{index + 1}_{sub_index + 1}] [STEP] [1] [글작성 토큰] [실패] [응답 : {response.status_code}] [내용 : {escaped_text}]')
                return None, None
        except Exception as e:
            self.log_updated.emit(f'[Index:{repeat + 1}_{index + 1}_{sub_index + 1}] [STEP] [1] 에러발생 : {e}')
            return None, None

    def check_post_client_token_for_cloud_flare(self, url, repeat, index, sub_index):
        try:
            host = get_base_domain(url)
            full_base_url = get_full_base_domain(url)
            post_client_token_url = f'{full_base_url}/set_post_client_token.cm'

            cookie_header = "; ".join([f"{cookie['name']}={cookie['value']}" for cookie in self.driver.get_cookies()])

            post_client_token_headers = {
                "Cookie": cookie_header,
                "accept": "application/json, text/javascript, */*; q=0.01",
                "accept-encoding": "gzip, deflate, br, zstd",
                "accept-language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
                "origin": full_base_url,
                "referer": url,
                "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Whale/3.28.266.14 Safari/537.36"
            }

            # 클라이언트를 통해 HTTP/2 요청 수행
            response = self.client.post(post_client_token_url, headers=post_client_token_headers, timeout=10)

            # 응답 내용 처리
            escaped_text = response.text.replace("<", "&lt;").replace(">", "&gt;")
            if response.status_code == 200:
                self.log_updated.emit(f'[Index:{repeat + 1}_{index + 1}_{sub_index + 1}] [STEP] [2] [유효성 검사] [응답 : {response.status_code}] [내용 : {escaped_text}]')
                return True
            else:
                self.log_updated.emit(f'[Index:{repeat + 1}_{index + 1}_{sub_index + 1}] [STEP] [2] [유효성 검사] [응답 : {response.status_code}] [내용 : {escaped_text}]')
                return False
        except Exception as e:
            self.log_updated.emit(f'[Index:{repeat + 1}_{index + 1}_{sub_index + 1}] [STEP] [2] 에러발생 : {e}')

    def post_add_for_cloud_flare(self, url, content, name, title, repeat, index, sub_index):
        try:
            sub_path = get_sub_path(url)
            menu_url = f'/{sub_path}/'
            board_code = get_board_value(url)
            host = get_base_domain(url)
            full_base_url = get_full_base_domain(url)
            post_url = f'{full_base_url}/backpg/post_add.cm'

            # # write_token 및 write_token_key 추출
            write_token = self.driver.find_element(By.CSS_SELECTOR, 'input[name="write_token"]').get_attribute("value")
            write_token_key = self.driver.find_element(By.CSS_SELECTOR, 'input[name="write_token_key"]').get_attribute(
                "value")

            cf_turnstile_response = self.driver.find_element(By.CSS_SELECTOR, 'input[name="cf-turnstile-response"]').get_attribute(
                "value")

            request_data = {
                'cf-turnstile-response': cf_turnstile_response,
                'idx': '0',
                'menu_url': menu_url,
                'back_url': '',
                'back_page_num': '',
                'board_code': board_code,
                'body': content,
                'plain_body': '',
                'is_editor': 'Y',
                'represent_img': '',
                'img': '',
                'img_tmp_no': '',
                'is_notice': 'no',
                'category_type': '0',
                'write_token': write_token,
                'write_token_key': write_token_key,
                'is_secret_post': 'no',
                'nick': name,
                'secret_pass': 'sdggs',
                'subject': title
            }

            cookie_header = "; ".join([f"{cookie['name']}={cookie['value']}" for cookie in self.driver.get_cookies()])
            encoded_payload = urllib.parse.urlencode(request_data)
            content_length = len(encoded_payload.encode("utf-8"))
            post_client_token_headers = {
                "Cookie": cookie_header,
                "accept": "application/json, text/javascript, */*; q=0.01",
                "accept-encoding": "gzip, deflate, br, zstd",
                "accept-language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
                "Referer": url,
                "Origin": full_base_url,
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Whale/3.28.266.14 Safari/537.36",
                "content-type": "application/x-www-form-urlencoded",
                "x-requested-with": "XMLHttpRequest",
                "sec-ch-ua": '"Whale";v="3", "Not-A.Brand";v="8", "Chromium";v="128"',
                "sec-ch-ua-mobile": "?0",
                "sec-ch-ua-platform": '"Windows"',
                "sec-fetch-dest": "empty",
                "sec-fetch-mode": "cors",
                "sec-fetch-site": "same-origin",
            }

            # 클라이언트를 통해 HTTP/2 요청 수행
            response = self.client.post(post_url, data=request_data, headers=post_client_token_headers, timeout=10)
            escaped_text = response.text.replace("<", "&lt;").replace(">", "&gt;")
            if response.status_code == 200:
                self.log_updated.emit(f'[Index:{repeat + 1}_{index + 1}_{sub_index + 1}] [STEP] [3] [글작성] [성공] [응답 : {response.status_code}] [내용 : {escaped_text}]')
                return True
            else:
                self.log_updated.emit(f'[Index:{repeat + 1}_{index + 1}_{sub_index + 1}] [STEP] [3] [글작성] [실패] [응답 : {response.status_code}] [내용 : {escaped_text}]')
                return False
        except Exception as e:
            self.log_updated.emit(f'[Index:{repeat + 1}_{index + 1}_{sub_index + 1}] [STEP] [3] 에러발생 : {e}')
            return False

    def init_driver(self):
        try:
            self.log_updated.emit('크롬 초기화 시작')
            chrome_options = Options()

            options_list = [
                f"user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/94.0.4606.61 Safari/537.36",
                "--disable-blink-features=AutomationControlled",
                "--start-maximized",
                "--log-level=3",
                "--disable-infobars",
                "--no-sandbox",
                "--disable-dev-shm-usage"
            ]

            for option in options_list:
                chrome_options.add_argument(option)

            chrome_options.add_argument("--disable-blink-features=AutomationControlled")
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_experimental_option("useAutomationExtension", False)

            self.driver = webdriver.Chrome(service=Service('./chromedriver.exe'), options=chrome_options)
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            self.log_updated.emit('크롬 초기화 완료')
        except Exception as e:
            print(e)

    def run(self):
        for k in range(self.repeat):
            try:
                if not self.use_chrome:
                    self.init_driver()
                self.log_updated.emit(f'{k + 1}번째 반복 실행')
                for i, entry in enumerate(self.entries):
                    try:
                        url, user_id, password = entry
                        for j in range(self.num_tabs):
                            perform_result = self.perform_task(k, i, j, url)

                        self.progress_updated.emit(i + 1)
                        if i + 1 != len(self.entries):
                            self.log_updated.emit(f'다음 글쓰기 까지 {self.writing_delay}초 대기')
                            time.sleep(self.writing_delay)

                    except Exception as e:
                        print_with_debug(e)
                self.log_updated.emit(f'다음 반복 까지 {self.overall_delay}초 대기')
                self.current_file_index = (self.current_file_index + 1) % len(self.excel_file_list)
                time.sleep(self.overall_delay)
            except Exception as e:
                print_with_debug(e)
        if self.driver:
            self.driver.quit()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.filename = 'urls.json'
        self.login_manage_button = QPushButton("로그인주소 관리")
        self.write_button = QPushButton("글쓰기 시작")
        self.title_modify_button = QPushButton("변경")
        self.add_button = QPushButton("추가")
        self.delete_button = QPushButton("삭제")
        self.table = QTableWidget()
        self.progress_bar = QProgressBar(self)
        self.log_edit_box = QTextEdit(self)
        self.error_log_edit_box = QTextEdit(self)
        self.count_label = None
        self.url_manager = None
        self.worker = None
        self.init_ui()
        self.load_settings()
        self.loadFilePaths()
        self.load_urls_from_file(self.filename)

    def init_ui(self):
        try:
            current_ver = load_version()
            self.setWindowTitle(f'SmartWriter - {current_ver}')
            self.setWindowIcon(QIcon('main_icon.png'))

            writing_delay_label = QLabel('글쓰기 딜레이(초)')
            self.writing_delay_input = QLineEdit()
            overall_delay_label = QLabel('전체 딜레이(초)')
            self.overall_delay_input = QLineEdit()
            repeat_label = QLabel('전체 반복 회수')
            self.repeat_input = QLineEdit()
            tab_label = QLabel('사이트당 게시물 작성 회수')
            self.tab_input = QLineEdit()
            x_pos_label = QLabel('클플 X 좌표')
            self.x_pos_input = QLineEdit()
            y_pos_label = QLabel('클플 Y 좌표')
            self.y_pos_input = QLineEdit()
            self.convert_checkbox = QCheckBox('특수문자 치환')
            self.use_proxy_checkbox = QCheckBox('프록시 사용')
            self.use_chrome_checkbox = QCheckBox('클라우드 플레어 우회')
            self.use_secret_checkbox = QCheckBox('크롬 프로필 복사')
            name_language_label = QLabel('이름 언어')
            self.name_language_combo_box = QComboBox()
            self.name_language_combo_box.addItems(['한글', '영어', '한글+숫자', '영어+숫자', '숫자만', '중국어', '일본어', '커스텀'])

            # Set the size policy for the new input boxes
            self.writing_delay_input.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            self.overall_delay_input.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            self.repeat_input.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            self.tab_input.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

            # Create a layout for the new input boxes and labels
            delay_layout = QHBoxLayout()
            delay_layout.addWidget(name_language_label)
            delay_layout.addWidget(self.name_language_combo_box)
            delay_layout.addWidget(writing_delay_label)
            delay_layout.addWidget(self.writing_delay_input)
            delay_layout.addWidget(overall_delay_label)
            delay_layout.addWidget(self.overall_delay_input)
            delay_layout.addWidget(repeat_label)
            delay_layout.addWidget(self.repeat_input)
            delay_layout.addWidget(tab_label)
            delay_layout.addWidget(self.tab_input)
            delay_layout.addWidget(x_pos_label)
            delay_layout.addWidget(self.x_pos_input)
            delay_layout.addWidget(y_pos_label)
            delay_layout.addWidget(self.y_pos_input)
            delay_layout.addWidget(self.convert_checkbox)
            delay_layout.addWidget(self.use_proxy_checkbox)
            delay_layout.addWidget(self.use_chrome_checkbox)

            self.load_from_file_button = QPushButton('파일에서 불러오기')
            self.load_from_file_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
            self.add_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
            self.delete_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

            self.url_edit_box = QLineEdit()
            self.id_edit_box = QLineEdit()
            self.pw_edit_box = QLineEdit()

            inner_layout = QVBoxLayout()
            inner_layout_top_layout = QHBoxLayout()
            inner_layout_top_layout.addWidget(QLabel('URL:'))
            inner_layout_top_layout.addWidget(self.url_edit_box, stretch=10)
            inner_layout_top_layout.addWidget(self.load_from_file_button, stretch=1)
            inner_layout_top_layout.addWidget(self.add_button, stretch=1)
            inner_layout_top_layout.addWidget(self.delete_button, stretch=1)

            inner_layout.addLayout(inner_layout_top_layout)
            self.url_table_widget = QTableWidget()
            self.url_table_widget.setColumnCount(1)
            self.url_table_widget.setHorizontalHeaderLabels(['URL'])
            self.url_table_widget.setSelectionBehavior(QTableWidget.SelectRows)
            self.url_table_widget.setSelectionMode(QTableWidget.ExtendedSelection)
            self.url_table_widget.setColumnWidth(0, 1200)
            self.url_table_widget.verticalHeader().setVisible(False)

            # 마지막에 이 코드를 추가하여 URL 열이 나머지 너비를 차지하도록 설정합니다.
            self.url_table_widget.horizontalHeader().setStretchLastSection(True)
            inner_layout.addWidget(self.url_table_widget)

            group_box = QGroupBox()
            group_box.setStyleSheet("QGroupBox { border: 1px solid gray; }")
            group_box.setLayout(inner_layout)

            self.load_from_file_button.clicked.connect(self.on_load_from_file_button_click)
            self.add_button.clicked.connect(self.on_add_button_click)
            self.delete_button.clicked.connect(self.on_delete_button_click)

            main_layout = QVBoxLayout()
            main_layout.addLayout(delay_layout)

            button_layout = QHBoxLayout()
            button_layout.addWidget(self.write_button)

            main_layout.addLayout(button_layout)

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
            self.error_log_edit_box.setReadOnly(True)  # 읽기 전용으로 설정
            self.error_log_edit_box.setFixedHeight(100)  # 높이 설정
            main_layout.addWidget(self.log_edit_box)
            main_layout.addWidget(self.error_log_edit_box)

            container = QWidget()
            container.setLayout(main_layout)
            self.setCentralWidget(container)
            self.resize(1600, 800)
            x, y = get_center_position(self.width(), self.height())
            self.move(x, y)

            self.check_for_update()
        except Exception as e:
            print_with_debug(e)

    def check_for_update(self):
        try:
            current_version = load_version()
            latest_version, download_url = get_latest_release()
            self.add_log(
                f'[프로그램 시작] [현재 버전 : {current_version}] [최신 버전 : {latest_version}]')

            if latest_version and download_url:
                if compare_versions(current_version, latest_version):
                    current_folder = os.getcwd()
                    dialog = UpdateDialog(current_version, latest_version)
                    update_accepted = dialog.get_update_decision()
                    if update_accepted:
                        self.perform_update(download_url, current_folder, latest_version)
                    else:
                        print("Update cancelled by the user.")

        except Exception as e:
            print_with_debug(e)

    def move_extracted_files(self, src_folder, dst_folder):
        try:
            for item in os.listdir(src_folder):
                src_path = os.path.join(src_folder, item)
                dst_path = os.path.join(dst_folder, item)
                if os.path.isfile(src_path):
                    shutil.move(src_path, dst_path)
                elif os.path.isdir(src_path):
                    for sub_item in os.listdir(src_path):
                        sub_src_path = os.path.join(src_path, sub_item)
                        sub_dst_path = os.path.join(dst_folder, sub_item)
                        if os.path.isfile(sub_src_path):
                            shutil.move(sub_src_path, sub_dst_path)
                        elif os.path.isdir(sub_src_path):
                            shutil.copytree(sub_src_path, sub_dst_path, dirs_exist_ok=True)
                    shutil.rmtree(src_path)
        except Exception as e:
            print_with_debug(e)

    def perform_update(self, download_url, current_folder, latest_version):
        temp_dir = os.path.join(os.path.expanduser("~"), "smartwriter_update")
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)

        self.add_log('새로운 업데이트 버전을 다운로드 합니다. 네트워크 상황에 따라 다소 시간이 걸릴 수 있습니다.')
        success = download_and_extract_zip(download_url, temp_dir)
        if not success:
            self.add_log(f'다운로드 실패({latest_version})')
            return

        self.add_log(f'다운로드 완료({latest_version})')

        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        new_version_folder = os.path.join(desktop_path, os.path.basename(current_folder) + "_" + latest_version)

        self.add_log(f'압축 해제 중. 해제 경로 : [{new_version_folder}]')

        if not os.path.exists(new_version_folder):
            os.makedirs(new_version_folder)

        # 압축 해제된 파일을 new_version_folder로 이동
        self.move_extracted_files(temp_dir, new_version_folder)

        self.add_log(f'압축 해제 완료({latest_version})')
        self.add_log(f'백업 파일 복사 시작')
        self.backup_files(current_folder, new_version_folder)
        self.add_log(f'백업 파일 복사 완료')

    def backup_files(self, src_folder, dst_folder):
        try:
            # 1. 모든 파일을 복사하되, .py와 .exe 확장자는 제외
            for item in os.listdir(src_folder):
                src_path = os.path.join(src_folder, item)
                dst_path = os.path.join(dst_folder, item)
                if os.path.isfile(src_path):
                    if not src_path.endswith(('.py', '.exe', 'version.json')):
                        shutil.copy2(src_path, dst_path)
                elif os.path.isdir(src_path) and item != '_internal':  # _internal 폴더는 제외
                    shutil.copytree(src_path, dst_path, dirs_exist_ok=True)

            # 2. _internal\txt 폴더의 모든 파일을 복사
            internal_src_folder = os.path.join(src_folder, '_internal', 'txt')
            internal_dst_folder = os.path.join(dst_folder, '_internal', 'txt')
            if not os.path.exists(internal_dst_folder):
                os.makedirs(internal_dst_folder)

            if os.path.exists(internal_src_folder):
                for item in os.listdir(internal_src_folder):
                    src_path = os.path.join(internal_src_folder, item)
                    dst_path = os.path.join(internal_dst_folder, item)
                    if os.path.isfile(src_path):
                        shutil.copy2(src_path, dst_path)
                    elif os.path.isdir(src_path):
                        shutil.copytree(src_path, dst_path, dirs_exist_ok=True)
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
            'repeat': self.repeat_input.text(),
            'tab': self.tab_input.text(),
            'xpos': self.x_pos_input.text(),
            'ypos': self.y_pos_input.text()
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
                self.tab_input.setText(settings.get('tab', ''))
                self.x_pos_input.setText(settings.get('xpos', ''))
                self.y_pos_input.setText(settings.get('ypos', ''))
        except Exception as e:
            print_with_debug(e)

    def add_log(self, log):
        if log.startswith('[INFO]'):
            log_box = self.error_log_edit_box
        else:
            log_box = self.log_edit_box

        current_time = datetime.datetime.now()
        formatted_time = current_time.strftime("[%Y-%m-%d %H:%M:%S]")
        formatted_log = f"{formatted_time} {log}"
        print(formatted_log)
        log_box.append(formatted_log)
        log_box.verticalScrollBar().setValue(log_box.verticalScrollBar().maximum())

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

    def on_write_button_click(self):
        self.save_settings()
        entries = self.get_all_entries()
        self.progress_bar.setMaximum(len(entries))
        self.worker = Worker(entries, int(self.writing_delay_input.text()), int(self.overall_delay_input.text()),
                             int(self.repeat_input.text()), self.convert_checkbox.isChecked(), self.getFileList(),
                             self.name_language_combo_box.currentText(), int(self.tab_input.text()), self.use_proxy_checkbox.isChecked(), self.use_chrome_checkbox.isChecked(),int(self.x_pos_input.text()), int(self.y_pos_input.text()) )
        self.worker.progress_updated.connect(self.update_progress)
        self.worker.log_updated.connect(self.add_log)
        self.worker.start()


    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def read_text_file(self, file_name):
        with open(file_name, 'r', encoding='utf-8') as file:
            lines = file.read().splitlines()
        return lines
    def on_load_from_file_button_click(self):
        try:
            options = QFileDialog.Options()
            options |= QFileDialog.DontUseNativeDialog
            file_name, _ = QFileDialog.getOpenFileName(self, "Open Text File", "", "Text Files (*.txt)",
                                                       options=options)

            if file_name:
                text_list = self.read_text_file(file_name)

                for url in text_list:
                    row_position = self.url_table_widget.rowCount()
                    self.url_table_widget.insertRow(row_position)
                    url_item = QTableWidgetItem(url)
                    url_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)  # Read-only 설정
                    self.url_table_widget.setItem(row_position, 0, url_item)

                self.write_urls_to_file(self.filename)

        except Exception as e:
            print_with_debug(e)

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


class UpdateDialog(QtWidgets.QDialog):
    def __init__(self, current_version, latest_version):
        super().__init__()
        self.current_version = current_version
        self.latest_version = latest_version
        self.update_accepted = False
        self.init_ui()
        self.center()

    def init_ui(self):
        self.setWindowTitle("신규 업데이트")
        self.setWindowIcon(QIcon('main_icon.png'))
        self.setGeometry(400, 400, 300, 150)
        self.label = QtWidgets.QLabel(f"현재 버전 : {self.current_version}\n새 업데이트 버전({self.latest_version}) 이 있습니다.", self)
        self.label.setGeometry(50, 20, 200, 30)

        self.update_button = QtWidgets.QPushButton("업데이트", self)
        self.update_button.setGeometry(50, 70, 80, 30)
        self.update_button.clicked.connect(self.accept_update)

        self.cancel_button = QtWidgets.QPushButton("취소", self)
        self.cancel_button.setGeometry(150, 70, 80, 30)
        self.cancel_button.clicked.connect(self.reject_update)

    def center(self):
        qr = self.frameGeometry()
        cp = QtWidgets.QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def accept_update(self):
        self.update_accepted = True
        self.close()

    def reject_update(self):
        self.update_accepted = False
        self.close()

    def get_update_decision(self):
        self.exec_()
        return self.update_accepted


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