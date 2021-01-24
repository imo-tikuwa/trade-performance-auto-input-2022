# -*- coding: utf-8 -*-
import os
import sys
# スクレイピング用
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
# ロギング用
import logging
import logzero
from logzero import logger
# スリープ用
import time
# 設定ファイル読み込み&出力
import configparser
# ディレクトリ、ファイルパス取得用
import tkinter, tkinter.filedialog
# コマンドラインパーサ
import click
# 現在日付取得用
from datetime import datetime
# Excelファイル操作
import win32com.client
from pywintypes import TimeType
# 暗号化/複合化を行うクラス
from encrypter import simple_encrypter
# バックアップ用
import shutil
# ランダム文字列生成用
import random
import string

# 定数
CONFIG_FILE_NAME = 'settings.ini'
CONFIG_OPT_LOGIN_ID = 'login_id'
CONFIG_OPT_PASSWORD = 'password'
CONFIG_OPT_CHROME_EXECUTABLE_PATH = 'chrome_executable_path'
CONFIG_OPT_ENCRYPTION_KEY = 'encryption_key'
CONFIG_OPT_TRADE_PERFORMANCE_XLSX_PATH = 'trade_performance_xlsx_path'
CURRENT_DATE = datetime.now().strftime("%Y%m%d")
WORK_DIR = 'work' + os.sep + CURRENT_DATE + os.sep
LOG_DIR = 'log' + os.sep
LOG_FILE = LOG_DIR + 'application.log'
# Trade-Performance(2021年度)の各月の入力行数
BUSINESS_DAY_EXCEL_ROW_MAP = {
    1: [4, 22],
    2: [4, 21],
    3: [4, 26],
    4: [4, 24],
    5: [4, 21],
    6: [4, 25],
    7: [4, 23],
    8: [4, 24],
    9: [4, 23],
    10: [4, 24],
    11: [4, 23],
    12: [4, 25],
}

# ログファイル
if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)
logzero.logfile(LOG_FILE, encoding = "utf-8")
logzero.loglevel(logging.INFO)

def get_config():
    """
    設定ファイルから設定を取得する
    設定ファイルに設定が存在しない場合は対話形式で設定を作成する
    """
    config = configparser.ConfigParser()
    config.read(CONFIG_FILE_NAME, 'cp932')
    config_section_name = 'default'

    # 設定ファイル内にchrome,defaultセクションが存在しない場合は作成
    if not config.has_section(config_section_name):
        logger.debug("{0}に{1}セクションを作成します".format(CONFIG_FILE_NAME, config_section_name))
        config.add_section(config_section_name)
    if not config.has_section(config_section_name):
        logger.debug("{0}に{1}セクションを作成します".format(CONFIG_FILE_NAME, config_section_name))
        config.add_section(config_section_name)

    # 設定ファイル内に暗号化に使用するキーが存在しない場合は作成
    if not config.has_option(config_section_name, CONFIG_OPT_ENCRYPTION_KEY):
        logger.debug("暗号化に使用するキーが存在しないため生成します")
        encryption_key = ''.join(random.choices(string.ascii_letters + string.digits, k=16))
        config.set(config_section_name, CONFIG_OPT_ENCRYPTION_KEY, encryption_key)
        with open(CONFIG_FILE_NAME, 'w') as config_file:
            logger.debug("{0}の{1}セクションに{2}を追記して保存します".format(CONFIG_FILE_NAME, config_section_name, CONFIG_OPT_ENCRYPTION_KEY))
            config.write(config_file)
    else:
        encryption_key = config.get(config_section_name, CONFIG_OPT_ENCRYPTION_KEY)

    # ChromeDriverのパス
    if not config.has_option(config_section_name, CONFIG_OPT_CHROME_EXECUTABLE_PATH):
        logger.info("ChromeDriverのパスが設定されていないので、設定してください")
        root = tkinter.Tk()
        root.withdraw()
        chrome_executable_path = tkinter.filedialog.askopenfilename(filetypes = [("ChromeDriverの実行ファイル", "*.exe")], initialdir = os.getcwd())
        if chrome_executable_path == '':
            logger.error("ChromeDriverのパスは必須です")
            sys.exit(1)
        logger.debug("選択されたファイル：{0}".format(chrome_executable_path))
        config.set(config_section_name, CONFIG_OPT_CHROME_EXECUTABLE_PATH, chrome_executable_path)
        with open(CONFIG_FILE_NAME, 'w') as config_file:
            logger.debug("{0}の{1}セクションに{2}を追記して保存します".format(CONFIG_FILE_NAME, config_section_name, CONFIG_OPT_CHROME_EXECUTABLE_PATH))
            config.write(config_file)
    else:
        chrome_executable_path = config.get(config_section_name, CONFIG_OPT_CHROME_EXECUTABLE_PATH)

    # ログインID
    if not config.has_option(config_section_name, CONFIG_OPT_LOGIN_ID):
        logger.info("SBI証券のログインIDを入力してください")
        login_id = input("入力：")
        if login_id == '':
            logger.error("SBI証券のログインIDは必須です")
            sys.exit(1)
        logger.debug("入力されたログインID：{0}".format(login_id))
        saving_login_id = simple_encrypter.encrypt(login_id, encryption_key)
        config.set(config_section_name, CONFIG_OPT_LOGIN_ID, saving_login_id)
        with open(CONFIG_FILE_NAME, 'w') as config_file:
            logger.debug("{0}の{1}セクションに{2}を追記して保存します".format(CONFIG_FILE_NAME, config_section_name, CONFIG_OPT_LOGIN_ID))
            config.write(config_file)
    else:
        login_id = config.get(config_section_name, CONFIG_OPT_LOGIN_ID)
        login_id = simple_encrypter.decrypt(login_id, encryption_key)

    # ログインパスワード(パスワードはsettings.iniに保存する際に暗号化する)
    if not config.has_option(config_section_name, CONFIG_OPT_PASSWORD):
        logger.info("SBI証券のログインパスワードを入力してください")
        password = input("入力：")
        if password == '':
            logger.error("SBI証券のログインパスワードは必須です")
            sys.exit(1)
        saving_password = simple_encrypter.encrypt(password, encryption_key)
        config.set(config_section_name, CONFIG_OPT_PASSWORD, saving_password)
        with open(CONFIG_FILE_NAME, 'w') as config_file:
            logger.debug("{0}の{1}セクションに{2}を追記して保存します".format(CONFIG_FILE_NAME, config_section_name, CONFIG_OPT_PASSWORD))
            config.write(config_file)
    else:
        password = config.get(config_section_name, CONFIG_OPT_PASSWORD)
        password = simple_encrypter.decrypt(password, encryption_key)

    # Trade-Performance-2021(xlsx)のパス
    if not config.has_option(config_section_name, CONFIG_OPT_TRADE_PERFORMANCE_XLSX_PATH):
        logger.info("Trade-Performance-2021.xlsxのパスが設定されていないので、設定してください")
        root = tkinter.Tk()
        root.withdraw()
        trade_performance_xlsx_path = tkinter.filedialog.askopenfilename(filetypes = [("Trade-Performance-2021.xlsx", "*.xlsx")], initialdir = os.getcwd())
        if trade_performance_xlsx_path == '':
            logger.error("Trade-Performance-2021.xlsxのパスは必須です")
            sys.exit(1)
        logger.debug("選択されたファイル：{0}".format(trade_performance_xlsx_path))
        config.set(config_section_name, CONFIG_OPT_TRADE_PERFORMANCE_XLSX_PATH, trade_performance_xlsx_path)
        with open(CONFIG_FILE_NAME, 'w') as config_file:
            logger.debug("{0}の{1}セクションに{2}を追記して保存します".format(CONFIG_FILE_NAME, config_section_name, CONFIG_OPT_TRADE_PERFORMANCE_XLSX_PATH))
            config.write(config_file)
    else:
        trade_performance_xlsx_path = config.get(config_section_name, CONFIG_OPT_TRADE_PERFORMANCE_XLSX_PATH)

    return {
        'chrome_executable_path': chrome_executable_path,
        'login_id': login_id,
        'password': password,
        'trade_performance_xlsx_path': trade_performance_xlsx_path,
        'encryption_key': encryption_key,
    }

def save_current_html_source(driver, debug_log_title, htmlname):
    """
    seleniumが参照中のhtmlソースを保存する
    """
    logger.debug(debug_log_title)
    with open(WORK_DIR + htmlname, 'w', encoding='utf-8') as f:
        f.write(driver.page_source)

@click.command(context_settings = dict(help_option_names = ['-h', '--help']))
@click.option('--debug', is_flag = True, help = "debugログを出力します")
def main(debug):

    logger.info("trade-performance-auto-input-2021 start.")

    if debug:
        logzero.loglevel(logging.DEBUG)

    # 設定取得
    config = get_config()

    logger.info("workディレクトリに本日分の作業フォルダ作成")
    if not os.path.exists(WORK_DIR):
        os.makedirs(WORK_DIR)

    logger.info("Chromeを起動")
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.104 Safari/537.36')
    options.add_argument('--guest')
    if not debug:
        options.add_argument('--headless')

    driver = webdriver.Chrome(executable_path = config['chrome_executable_path'], options = options)

    logger.info("SBI証券にログイン")
    logger.debug("ログイン画面を開く")
    driver.get("https://www.sbisec.co.jp/ETGate")

    logger.debug("ログイン情報を入力、ログインボタンクリック")
    driver.find_element_by_css_selector("input[name='user_id']").send_keys(config['login_id'])
    driver.find_element_by_css_selector("input[name='user_password']").send_keys(config['password'])
    driver.find_element_by_css_selector("input[name='ACT_login']").click()
    if debug:
        save_current_html_source(driver, 'ログイン後画面のソースを保存', 'login_after.html')

    logger.info("口座管理画面を開く")
    driver.get('https://site2.sbisec.co.jp/ETGate/?_ControlID=WPLETacR001Control&_PageID=DefaultPID&_DataStoreID=DSWPLETacR001Control&_ActionID=DefaultAID&getFlg=on')
    if debug:
        save_current_html_source(driver, '口座管理画面のソースを保存', 'account_management.html')

    logger.debug('口座管理画面の「計」を取得')
    sum_selector_path = 'body > div:nth-child(1) > table > tbody > tr > td:nth-child(1) > table > tbody > tr:nth-child(2) > td > table:nth-child(1) > tbody > tr > td > form > table:nth-child(3) > tbody > tr:nth-child(1) > td:nth-child(2) > table:nth-child(20) > tbody > tr > td:nth-child(1) > table:nth-child(7) > tbody > tr:nth-child(8) > td:nth-child(2) > div > b'
    current_sum = driver.find_element_by_css_selector(sum_selector_path).text

    logger.info("GoogleChrome正常終了")
    driver.close()
    driver.quit()

    logger.info("Excelファイルのバックアップを作成")
    shutil.copy(config['trade_performance_xlsx_path'], WORK_DIR + os.path.basename(config['trade_performance_xlsx_path']))
    today_md_slash = datetime.today().strftime('%#m/%#d')
    today_m_int = int(datetime.today().strftime('%#m'))

    logger.debug("pywin32でExcelファイルを開く")
    app = win32com.client.Dispatch("Excel.Application")
    wb = app.Workbooks.Open(config['trade_performance_xlsx_path'])

    target_sheet_name = datetime.today().strftime('%#m') + '月'
    logger.debug(target_sheet_name + 'のシート取得')
    ws = wb.Worksheets(target_sheet_name)

    # A列の月日とプログラムの実行日を比較して↑で取得した「計」の書き込み先を見つける
    target_row_num = None
    for row_num in range(BUSINESS_DAY_EXCEL_ROW_MAP[today_m_int][0], BUSINESS_DAY_EXCEL_ROW_MAP[today_m_int][1]):
        cell_value = ws.Range('A' + str(row_num)).Value
        if (type(cell_value) is TimeType and today_md_slash == cell_value.strftime('%#m/%#d')):
            logger.debug('A列に今日の日付が見つかりました。見つけた日付：{0}'.format(today_md_slash))
            target_row_num = row_num
            break
        else:
            continue
        break

    if target_row_num is None:
        logger.debug('A列に今日の日付が見つかりませんでした。')
        wb.Close()
        app.Quit()
        sys.exit(1)

    # M列(口座A欄)に上で取得した計を追記して保存
    ws.Range('M' + str(target_row_num)).Value = current_sum
    logger.info('Excelファイルに追記しました。')
    wb.Save()
    logger.info('Excelファイルを上書き保存しました')
    wb.Close()
    app.Quit()

    logger.info("trade-performance-auto-input-2021 end.")
    return


if __name__ == "__main__":
    main()