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
from selenium.common.exceptions import NoSuchElementException, SessionNotCreatedException
# ロギング用
import logging
import logzero
from logzero import logger
# 設定ファイル読み込み&出力
import configparser
# ディレクトリ、ファイルパス取得用
import tkinter, tkinter.filedialog
# コマンドラインパーサ
import click
# 現在日時取得用
from datetime import datetime, timedelta, timezone
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

def is_before_trading():
    """
    取引時間前かどうかを判定する
    現在時刻が0時～9時のときtrueとなる
    """
    return 0 < CURRENT_HMS_INT < 90000

# プログラムで共通して使用するタイムゾーン生成、現在時刻取得
JST = timezone(timedelta(hours=+9), 'JST')
CURRENT_DATE = datetime.now(JST)
CURRENT_HMS_INT = int(CURRENT_DATE.strftime("%H%M%S"))
TARGET_MD_SLASH = (CURRENT_DATE - timedelta(days=1)).strftime('%#m/%#d') if is_before_trading() else CURRENT_DATE.strftime('%#m/%#d')
TARGET_M_INT = int((CURRENT_DATE - timedelta(days=1)).strftime('%#m')) if is_before_trading() else int(CURRENT_DATE.strftime('%#m'))

# 定数
CONFIG_FILE_NAME = 'settings.ini'
CONFIG_OPT_LOGIN_ID = 'login_id'
CONFIG_OPT_PASSWORD = 'password'
CONFIG_OPT_CHROME_EXECUTABLE_PATH = 'chrome_executable_path'
CONFIG_OPT_ENCRYPTION_KEY = 'encryption_key'
CONFIG_OPT_TRADE_PERFORMANCE_XLSX_PATH = 'trade_performance_xlsx_path'
if is_before_trading():
    WORK_DIR = 'work' + os.sep + (CURRENT_DATE - timedelta(days=1)).strftime("%Y%m%d") + os.sep
else:
    WORK_DIR = 'work' + os.sep + CURRENT_DATE.strftime("%Y%m%d") + os.sep
LOG_DIR = 'log' + os.sep
LOG_FILE = LOG_DIR + 'application.log'
# Trade-Performance(2022年度)の各月の入力行数
BUSINESS_DAY_EXCEL_ROW_MAP = {
    1: [4, 23],
    2: [4, 22],
    3: [4, 26],
    4: [4, 24],
    5: [4, 23],
    6: [4, 26],
    7: [4, 24],
    8: [4, 26],
    9: [4, 24],
    10: [4, 24],
    11: [4, 24],
    12: [4, 26],
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

    # Trade-Performance-2022(xlsx)のパス
    if not config.has_option(config_section_name, CONFIG_OPT_TRADE_PERFORMANCE_XLSX_PATH):
        logger.info("Trade-Performance-2022.xlsxのパスが設定されていないので、設定してください")
        root = tkinter.Tk()
        root.withdraw()
        trade_performance_xlsx_path = tkinter.filedialog.askopenfilename(filetypes = [("Trade-Performance-2022.xlsx", "*.xlsx")], initialdir = os.getcwd())
        if trade_performance_xlsx_path == '':
            logger.error("Trade-Performance-2022.xlsxのパスは必須です")
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


@click.command(context_settings = dict(help_option_names = ['-h', '--help']))
@click.option('--debug', is_flag = True, help = "debugログを出力します")
def main(debug):

    logger.info("trade-performance-auto-input-2022 start.")

    if debug:
        logzero.loglevel(logging.DEBUG)

    # 設定取得
    config = get_config()

    logger.info("workディレクトリに本日分の作業フォルダ作成")
    if not os.path.exists(WORK_DIR):
        os.makedirs(WORK_DIR)

    logger.info("Chromeを起動")
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation'])
    options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.104 Safari/537.36')
    options.add_argument('--guest')
    if not debug:
        options.add_argument('--headless')

    try:
        driver = webdriver.Chrome(executable_path = config['chrome_executable_path'], options = options)
    except SessionNotCreatedException as e:
        logger.error(e.msg)
        sys.exit(1)

    logger.info("SBI証券にログイン、口座管理画面を開き、資産合計を取得する")
    current_sum = None
    for trial_num in range(3):
        # ログイン後の口座管理画面のURLにアクセス
        driver.get('https://site2.sbisec.co.jp/ETGate/?_ControlID=WPLETacR001Control')
        logger.debug('{0}回目のアクセス'.format(trial_num + 1))
        try:
            # ログインセッションが存在しない場合ログイン画面が表示されるのでログイン情報を入力してログインボタンをクリック
            driver.find_element_by_css_selector("input[name='user_id']").send_keys(config['login_id'])
            driver.find_element_by_css_selector("input[name='user_password']").send_keys(config['password'])
            logger.debug('ログイン処理実行')
            driver.find_element_by_css_selector("input[name='ACT_login']").click()
        except NoSuchElementException as e:
            # 口座管理画面が開けた場合 = ログイン情報の入力欄が見つからずにNoSuchElementExceptionがスローされる
            # 口座管理画面から目的の値を取得する
            try:
                logger.debug('口座管理画面から保有資産の合計金額を取得')
                # cssセレクタについて信用取引の持越しがある/なしなんかで、HTMLの構造がちょっと変わった模様(no such elementの例外出てた)
                # そのためtrのWebElementリストを取得、ループして先頭のtdのテキストが「計」の行から、保有資産の合計金額を取得する形に変更
                tr_selector_path = 'body > div:nth-child(1) > table > tbody > tr > td:nth-child(1) > table > tbody > tr:nth-child(2) > td > table:nth-child(1) > tbody > tr > td > form > table:nth-child(3) > tbody > tr:nth-child(1) > td:nth-child(2) > table:nth-child(19) > tbody > tr > td:nth-child(1) > table:nth-child(7) > tbody > tr:nth-child(8)'
                for element in driver.find_elements_by_css_selector(tr_selector_path):
                    if element.find_element_by_css_selector('td:nth-child(1)').text == '計':
                        current_sum = element.find_element_by_css_selector('td:nth-child(2) > div > b').text
                        break
                break
            except NoSuchElementException as e:
                logger.error('口座管理画面から保有資産の合計金額の取得に失敗しました。')
                logger.error(e)
                sys.exit(1)

    # 目的の値が見つからないときに複数回口座管理画面を開く処理を実施、保有資産の合計金額が取得できなかったらエラー吐いて終了
    if current_sum is None:
        logger.error('口座管理画面から保有資産の合計金額の取得に失敗しました。')
        sys.exit(1)

    logger.info("GoogleChrome正常終了")
    driver.close()
    driver.quit()

    logger.info("Excelファイルのバックアップを作成")
    shutil.copy(config['trade_performance_xlsx_path'], WORK_DIR + os.path.basename(config['trade_performance_xlsx_path']))

    logger.debug("pywin32でExcelファイルを開く")
    app = win32com.client.Dispatch("Excel.Application")
    wb = app.Workbooks.Open(config['trade_performance_xlsx_path'])

    target_sheet_name = CURRENT_DATE.strftime('%#m') + '月'
    logger.debug(target_sheet_name + 'のシート取得')
    ws = wb.Worksheets(target_sheet_name)

    # A列の月日とプログラムの実行日を比較して↑で取得した「計」の書き込み先を見つける
    target_row_num = None
    for row_num in range(BUSINESS_DAY_EXCEL_ROW_MAP[TARGET_M_INT][0], BUSINESS_DAY_EXCEL_ROW_MAP[TARGET_M_INT][1]):
        cell_value = ws.Range('A' + str(row_num)).Value
        if (type(cell_value) is TimeType and TARGET_MD_SLASH == cell_value.strftime('%#m/%#d')):
            logger.debug('A列に今日の日付が見つかりました。見つけた日付：{0}'.format(TARGET_MD_SLASH))
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

    logger.info("trade-performance-auto-input-2022 end.")
    return


if __name__ == "__main__":
    main()