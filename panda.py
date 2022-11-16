from pickle import FALSE, NONE
from tkinter.constants import TRUE
from attr import NOTHING
from selenium import webdriver
from webdriver_manager import driver
from webdriver_manager.chrome import ChromeDriverManager  # chomewebdriverの自動更新ライブラリ
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from logmanager import logManager as LM
from panda_excel_manager import excelData
import time
import os
import glob
import sys
import re
import json
import shutil
from panda_value import (DLFOLDERNAME, addUpdateMsIDs, clearUpdateMsIDS,isConnectedDLPath,isConnected, DLROOTPATH, DOWNLOADDIRPATH, PROMOTIONSYS_MSURL, 
PROMOTIONSYS_LOGINURL, PROMOTIONSYS_ID, PROMOTIONSYS_PW, PROMOTIONSYS_DLTEXT, SYSDOWNLOADDIRPATH, Errors)
from panda_json import *


# グローバル変数
global_ms_json = {}  # プロモーション管理システムの原稿リストjson.実行時に最初に取得
download_enable = True  # ダウンロード実行可否：FalseならDLせずに、更新できる。
global_ed = excelData()  # Excelデータログ用
# グローバル変数ここまで

##関数定義##
#クリックのjavascript実行版：こっちの方が安定する
def clickWithJavascript(driver, element):
    driver.execute_script('arguments[0].click();', element)

# グループページのファイル履歴表示ボタンから、グループが更新されたかを返す：更新情報辞書（更新日時、更新者） or False
# ファイル履歴表示ボタン：id=history-g116（グループID）。all（まとめ）はここではチェックしない。
def IsUpdateMsWithOpenHistoryAll(driver, group_id, prev_update):
    isupdate = False
    updateinfo = {}
    if not group_id == "all":  # まとめは見なくてOK
        # 履歴取得
        try:
            history_elements = driver.find_elements(
                By.ID, "history-" + group_id)
            if (len(history_elements) > 0):  # IDだから複数あることは考えられない
                driver.execute_script(
                    "window.open('manuscript.php?mode=fh&type=" + group_id + "');")
                driver.switch_to.window(
                    driver.window_handles[-1])  # 履歴ウィンドウに移動
                # javascriptのテープル出力まで待つ。
                WebDriverWait(driver, 30).until(
                    EC.visibility_of_element_located((By.CLASS_NAME, 'bootstrap-table')))
                #該当するレコードがありません表示なら。エラーキャッチして続ける。
                WebDriverWait(driver, 30).until(
                    EC.invisibility_of_element_located((By.CLASS_NAME, 'no-records-found')))
                current_ms_update = driver.find_elements(
                    By.XPATH, '//table[@id="file-history-table"]/tbody/tr[1]/td')
                # 一応もう一回チェック
                if(len(current_ms_update) > 0):
                    # 更新時間がある
                    LM.loginfo("更新時間テーブルあり")
                    if(current_ms_update[0].text != prev_update):
                        updateinfo["update_time"] = current_ms_update[0].text #更新日時
                        updateinfo["author"] = current_ms_update[3].text #更新者
                        LM.loginfo("更新された" + "{" + prev_update + "}" + "→" + "{" + updateinfo["update_time"] + "}")
                        isupdate = True
                    else:
                        LM.loginfo("更新されてない")
                        isupdate = False
                else:
                    LM.loginfo("更新時間テーブルがない")
                    isupdate = False
        except Exception as e:
            LM.loginfo(str(e))
            LM.loginfo('【更新履歴がまだありません。】続けます。')

        # 履歴ウィンドウ閉じる
        driver.close()
        # 原稿ページに戻る
        driver.switch_to.window(driver.window_handles[0])
        if isupdate:
            return(updateinfo)
        else:
            return(isupdate)

# 管理システムからグループの最新更新日時を取得。""か最新の更新日を返す。
def getGroupUpdateTime(diver, group_id):
    output = ""
    if not group_id == "all":  # まとめは見なくてOK
        # 履歴取得
        try:
            history_elements = driver.find_elements(
                By.ID, "history-" + group_id)
            if (len(history_elements) > 0):  # IDだから複数あることは考えられない
                driver.execute_script(
                    "window.open('manuscript.php?mode=fh&type=" + group_id + "');")
                driver.switch_to.window(
                    driver.window_handles[-1])  # 履歴ウィンドウに移動
                # javascriptのテープル出力まで待つ：タイムアウトしたら履歴がない
                WebDriverWait(driver, 15).until(
                    EC.invisibility_of_element_located((By.CLASS_NAME, 'no-records-found')))
                update_time = driver.find_elements(
                    By.XPATH, '//table[@id="file-history-table"]/tbody/tr[1]/td')
                if(len(update_time) > 0):
                    # 更新時間がある
                    LM.loginfo("更新時間テーブルあり")
                    output = update_time[0].text
                else:
                    LM.loginfo("更新時間テーブルがない")
                    output = ""
            else:
                LM.loginfo("ファイル履歴表示ボタンがない")

        except Exception as e:
            LM.loginfo(e)
            LM.loginfo('【履歴エラー】')
        # 履歴ウィンドウ閉じる
        driver.close()
        # 原稿ページに戻る
        driver.switch_to.window(driver.window_handles[0])
        return(output)

# downloadFilePath内の全ファイルから最新ファイルの絶対パスを取得：.*、.tmp以外、.xlsmで終わる。対象ファイルがないとFalse
def getLatestDownloadedFileName(downloadFilePath):
    if len(os.listdir(downloadFilePath)) == 0:
        return False
    files = [downloadFilePath + "/" +
             f for f in os.listdir(downloadFilePath) if (not f.startswith('.')) and (not f.endswith('.tmp') and (f.endswith('.xlsm')))]
    if len(files) > 0:
        return os.path.abspath(max(files,key=os.path.getctime))
    else:
        return False

#downloadFilePath内の不要ファイルを全て削除：.xlsmと.crdownload
def removeFilesFromDownloadFolder(downloadFilePath):
    if len(os.listdir(downloadFilePath)) == 0:
        return False
    files = [downloadFilePath + "/" +
             f for f in os.listdir(downloadFilePath) if (f.endswith('.xlsm') or f.endswith('.crdownload'))]
    if len(files) > 0:
        for rmf in files:
            os.remove(os.path.abspath(rmf))
        return True
    else:
        return False


# 指定フォルダに、ファイルのダウンロードが完了するまで待機する
def wait_file_download(path, timeout):
    # 待機タイムアウト時間(秒)設定
    timeout_second = timeout

    # 指定時間分待機
    for i in range(timeout_second + 1):
        # ファイル一覧取得
        match_file_path = os.path.join(path, '*.*')
        files = glob.glob(match_file_path)

        # ファイルが存在する場合
        if files:
            # ファイル名の拡張子に、'.crdownload'が含むかを確認
            extensions = [
                file_name for file_name in files if '.crdownload' in os.path.splitext(file_name)]

            # '.crdownload'が見つからなかったら抜ける
            if not extensions:
                break

        # 指定時間待っても .crdownload 以外のファイルが確認できない場合 エラー
        if i >= timeout_second:
            # 終了処理
            raise Exception('file cannnot be finished downloading!')

        # 一秒待つ
        time.sleep(1)

    return


# fpathのファイルの存在が確認できるまで待つ
def wait_file_exist(fpath, timeout):
    # 待機タイムアウト時間(秒)設定
    timeout_second = timeout

    # 指定時間分待機
    for i in range(timeout_second + 1):

        # ファイルが存在する場合抜ける
        if os.path.exists(fpath):
            break

        # 指定時間待ってもファイルが確認できない場合 エラー
        if i >= timeout_second:
            # 終了処理
            raise Exception('file is not exist!')

        # 一秒待つ
        time.sleep(1)

    return

# フォルダが存在しなければ作成
def my_makedirs(path):
    if not os.path.isdir(path):
        os.makedirs(path)
        return True
    else:
        return False

# ChromeWebDriverの初期化：driver変数を返す
def initChromWebDriver(downloadPath):
    # Chromeドライバーオプション

    my_makedirs(downloadPath)

    LM.loginfo("download:"+downloadPath)
    options = webdriver.ChromeOptions()
    options.add_experimental_option("prefs", {"download.default_directory": downloadPath,
                                              "download.directory_upgrade": True})
    options.add_argument('--headless')  # ヘッドレスモード

    # chomewebdriverの自動更新
    chromewebdriver = webdriver.Chrome(
        ChromeDriverManager().install(), options=options)
        
    chromewebdriver.command_executor._commands["send_command"] = (
        "POST",
        '/session/$sessionId/chromium/send_command'
    )

    params = {
        'cmd': 'Page.setDownloadBehavior',
        'params': {
            'behavior': 'allow',
            'downloadPath': downloadPath
        }
    }
    chromewebdriver.execute("send_command", params=params)

    return chromewebdriver

# プロモーション管理システムにログイン
def loginSys(driver):

    # ログイン
    try:
        driver.get(PROMOTIONSYS_LOGINURL)
        WebDriverWait(driver, 30).until(EC.presence_of_all_elements_located)  # DOM要素読み込み待ち
        username = driver.find_element(By.ID, 'username')  # ユーザー名要素
        password = driver.find_element(By.ID, 'inputPassword')  # パスワード要素
        loginbtn = driver.find_element(By.XPATH, '//button[text()="Sign in"]')  # ログインボタン（指定方法要検討）
        username.send_keys(PROMOTIONSYS_ID)
        password.send_keys(PROMOTIONSYS_PW)
        # ログインクリック
        clickWithJavascript(driver, loginbtn) 
        #loginbtn.click()
        # ログインここまで
    except Exception as e:
        LM.loginfo(str(e))
        LM.loginfo('ログインエラー')
        return False

    return True
        
# 原稿データの初期化や更新(新規追加と削除) True：成功 False：失敗
def updateMsdata(driver):
    LM.loginfo("原稿データ更新確認開始")
    # システムから全原稿のIDをkeyとした、name、updatetimeのjsonオブジェクト取得
    ms_data = global_ms_json
    # jsonデータロード
    loaded_ms_data = INSTANCE_OF_PANDAJSON.getJsonData()
    # global_msdata.dumpsJsonData()
    for ms_id in ms_data:
        # IDがロードしたjsonデータになければ新規追加（update_timeは空）
        if not ms_id in loaded_ms_data:
            loaded_ms_data[ms_id] = {
                "name": ms_data[ms_id]["name"], "update_time": "", "error": [], "group": {}}
            LM.loginfo(ms_id + "は追加された")
    # jsonデータのIDがシステムの原稿IDリストになければ削除
    for ms_id in list(loaded_ms_data):  # 削除するときはlist化
        if not ms_id in ms_data:
            LM.loginfo(ms_id + "は削除された")
            del loaded_ms_data[ms_id]
    # nameの更新確認
    for ms_id in ms_data:
        # nameが違ったら更新。フォルダを作成。
        if ms_data[ms_id]["name"] != loaded_ms_data[ms_id]["name"]:
            ms_title_before = re.sub(r'[\\/:*?"<>|]+', '／', loaded_ms_data[ms_id]["name"])
            ms_title_after = re.sub(r'[\\/:*?"<>|]+', '／', ms_data[ms_id]["name"])

            #DLフォルダループ-START
            for dlpath in DOWNLOADDIRPATH:
                ms_path = os.path.join(dlpath, ms_title_after)
                text_path = os.path.join(ms_path, "タイトルが変更された.txt")
                # サーバー接続確認
                if isConnected(dlpath):
                    #フォルダ作成
                    if my_makedirs(ms_path):
                        #ログ出力
                        LM.loginfo("【"+ dlpath + "】" + ms_id + "のフォルダを作成しました。" + "{" + loaded_ms_data[ms_id]["name"] + "}")
                else:
                    LM.loginfo("【panda>updateMsdata】" + dlpath + "に接続されていない")
                    #一つでも接続できてなかったら止める
                    return False

                #テキストファイルで原稿が引き継がれている事を明示
                f = open(text_path, 'w')
                f.write('前のタイトル\n' + ms_title_before)
                f.close()
            #DLフォルダループ-END

            LM.loginfo(ms_id + "はタイトル変更された。" + "{" + ms_title_before + "}→{" + ms_title_after + "}")
            loaded_ms_data[ms_id]["name"] = ms_data[ms_id]["name"]

    # msdata.jsonに書き出し
    INSTANCE_OF_PANDAJSON.setDumpMsjson(loaded_ms_data)
    # INSTANCE_OF_PANDAJSON.dumpsJsonData()
    LM.loginfo("msdata.jsonアップデート完了（原稿）")
    return True

# 原稿データ内のグループデータをアップデート
def updateGroupdata(driver):
    LM.loginfo("グループ更新確認開始")
    # jsonデータロード
    loaded_ms_data = INSTANCE_OF_PANDAJSON.getJsonData()
    # INSTANCE_OF_PANDAJSON.dumpsJsonData()
    # 原稿データが無ければエラー終了
    if loaded_ms_data == {}:
        LM.loginfo("原稿データがありません。")
        driver.close()
        driver.quit()
        return
    else:
        # グループチェック
        for msid in loaded_ms_data:
            # idの原稿ページへ
            ms_page_url = PROMOTIONSYS_MSURL + "?mode=flier-select&id=" + msid
            try:
                # idの原稿ページURL
                driver.get(ms_page_url)
                # DOM要素読み込み待ち
                WebDriverWait(driver, 30).until(
                    EC.presence_of_all_elements_located)
                # タブ巡回:タブ要素があれば
                if(len(driver.find_elements(By.CLASS_NAME, 'group-tab')) > 0):
                    ms_group_tabs = driver.find_elements(
                        By.CLASS_NAME, 'group-tab')
                    for i in range(len(ms_group_tabs)):
                        g_tab_element = ms_group_tabs[i].find_element(
                            By.TAG_NAME, 'a')
                        ms_group_name = g_tab_element.text
                        ms_group_id = g_tab_element.get_attribute("data-key")
                        LM.loginfo(
                            "今{" + ms_group_id + ":" + ms_group_name + "}見てる")
                        # グループIDが原稿内のjsonデータになければ追加
                        # まとめグループは見ない
                        if not ms_group_id == "all":
                            if not ms_group_id in loaded_ms_data[msid]["group"]:
                                LM.loginfo(ms_group_id + "を追加")
                                loaded_ms_data[msid]["group"][ms_group_id] = {
                                    "name": ms_group_name, "update_time": "", }
                        else:
                            LM.loginfo("まとめスルー")
                else:
                    LM.loginfo("グループタブがない")

            except Exception as e:
                LM.loginfo(str(e))
                LM.loginfo('グループ更新エラー')
    # msdata.jsonに書き出し
    INSTANCE_OF_PANDAJSON.setDumpMsjson(loaded_ms_data)
    # INSTANCE_OF_PANDAJSON.dumpsJsonData()
    LM.loginfo("msdata.jsonアップデート完了（グループ）")

# 原稿IDをkeyとしたJsonオブジェクトのリストをプロモーション管理システムから取得。Jsonオブジェクト返す。エラーはFalse返す。
def getMsOBJListFromSys(driver):
    msObj = {}  # 返り値用
    # INSTANCE_OF_PANDAJSON.dumpsJsonData()
    # 原稿確認ページへ
    driver.get(PROMOTIONSYS_MSURL)
    WebDriverWait(driver, 30).until(EC.presence_of_all_elements_located)  # DOM要素読み込み待ち
    # 自社担当分にする
    try:
        # 自社担当分ボタン
        clickWithJavascript(driver, driver.find_element(By.XPATH, '//div[@id="relation"]/label[1]')) 
        WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.CLASS_NAME, 'fixed-table-loading')))  # リストテープルが表示されるまで待つ
    except Exception as e:
        LM.loginfo(str(e))
        LM.loginfo('自社担当分エラー')
        return False
    #期間を2ヶ月先までに設定
    try:
        day_delta = 31 #終期を進める日数      
        #終了日を設定
        ms_span_to = driver.find_element(By.ID, 'date-to') #期間終了日要素
        ms_span_to.click() #カレンダー出す
        WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.CLASS_NAME, 'datepicker')))#カレンダー表示待ち
        #設定する年月まで進める。calendar_date（年月）が表示されるまで進める。
        cal_element = driver.switch_to.active_element
        for i in range(0, day_delta):
            cal_element.send_keys(Keys.ARROW_RIGHT)
        #カレンダーの日付選択
        cal_element.send_keys(Keys.ENTER)
        WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.CLASS_NAME, 'fixed-table-loading')))
    except Exception as e:
        #LM.loginfo(str(e))
        LM.loginfo('期間設定エラー')
        return False
        #driver.close()
        #driver.quit()
    # チェック開始
    try:
        WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.CLASS_NAME, 'fixed-table-loading')))  # リストテープルが表示されるまで待つ
        ms_pages = driver.find_elements(By.CLASS_NAME, 'page-number')
        data_index = 0

        for i in range(len(ms_pages)):
            ms_pages = driver.find_elements(By.CLASS_NAME, 'page-number')  # ページ遷移後はもう一回取得必要
            if i > 0:
                ms_pages[i].find_element(By.TAG_NAME, 'a').click()
            #テーブル表示待ち
            WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.CLASS_NAME, 'fixed-table-loading')))
            WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//table[@id="ms-manager-list"]/tbody/tr')))
            ms_managers = driver.find_elements(By.XPATH, '//table[@id="ms-manager-list"]/tbody/tr')
            #LM.loginfo(len(ms_managers))
            for j in range(len(ms_managers)):
                tds = ms_managers[j].find_elements(By.TAG_NAME, 'td')
                btn_ms = tds[7].find_element(By.LINK_TEXT, '原稿')
                btn_href = btn_ms.get_attribute('href')  # リンクURL
                # idをリンクURLから抽出、リストで返される
                ms_id = re.findall('.+id=(.+)', btn_href)
                ms_name = tds[2].text
                ms_update = tds[3].text
                #LM.loginfo(ms_id[0] +":"+ ms_name +":"+ ms_update)
                # jsonオブジェクト作成
                msObj[ms_id[0]] = {"name": ms_name, "update_time": ms_update, }
                data_index = data_index + 1
    except Exception as e:
        LM.loginfo(str(e))
        LM.loginfo('原稿更新エラー')
        return False

    json.dumps(msObj, ensure_ascii=False, indent=4)
    LM.loginfo("プロモーション管理システムの原稿リスト\n" + str(msObj))
    return msObj

# スクレイピング本体。更新されたms_idとgroup_idを返すの辞書を記録。エラーステートを返す。値はpanda.valueで管理。
def panda():
    clearUpdateMsIDS()  # 前回更新分をクリア

    ###メイン##
    try:
        # chromeDriver初期化：ダウンロードフォルダと原稿ダウンロードフォルダがなければ作成
        for try_count in range(3): # 最大3回ドライバー初期化トライ
            try:
                driver = initChromWebDriver(SYSDOWNLOADDIRPATH)
                
                #DLフォルダループ-START
                for rootpath in DLROOTPATH:
                    #サーバー接続確認
                    if(isConnected(rootpath)):
                        #原稿ダウンロードフォルダなければ作成
                        my_makedirs(os.path.join(rootpath, DLFOLDERNAME))
                    else:
                        LM.loginfo("【panda:driver初期化】" + rootpath + "に接続されていない。")
                        driver.close()
                        driver.quit()
                        #エラー：DLフォルダ未接続
                        return Errors.DLFOLDERCONNECTION.name
                #DLフォルダループ-END

            except Exception as e:
                LM.loginfo(str(e))
                LM.loginfo('【ChromeWebDriverの初期化に失敗】' + "{" + try_count + "}")
            else:
                #成功なら抜ける
                break
        else:
            #全部失敗したら終了
            LM.loginfo('【最大回数初期化に失敗】')
            #エラー：ドライバー初期化失敗
            return Errors.INITWEBDRIVER.name
        
        # ログイン
        if not loginSys(driver):
            driver.close()
            driver.quit()
            return Errors.LOGIN.name

        # msdata.json読み込み
        INSTANCE_OF_PANDAJSON.setjsonDataFromJson()
        # システムの原稿リスト取得
        global global_ms_json  # グローバル宣言
        global_ms_json = getMsOBJListFromSys(driver)
        if global_ms_json == False:
            #Falseはサーバー接続できてないので終了。Trueならそのまま続ける
            driver.close()
            driver.quit()
            #エラー：原稿エラー
            return Errors.MISSINGMSDATA.name
        
        # システムとmsdata.jsonの情報を比較して同期させる
        # global_ms_jsonとINSTANCE_OF_PANDAJSONを比較して、INSTANCE_OF_PANDAJSONとmsdata.jsonを更新。（追加・削除）。
        # 同じIDでタイトルのみ変更される場合ある。→nameを更新。フォルダ作成される。
        if not updateMsdata(driver):
            #Falseはサーバー接続できてないので終了。Trueならそのまま続ける
            driver.close()
            driver.quit()
            #エラー：DLフォルダ未接続
            return Errors.DLFOLDERCONNECTION.name
        
        global global_ed
        global_ed = excelData()  # Excelデータログ
        # グループリストを更新：必要ない？
        # updateGroupdata(driver)

        LM.loginfo("原稿データダウンロード開始")
        # jsonデータロード
        loaded_ms_data = INSTANCE_OF_PANDAJSON.getJsonData()

        # この時点で原稿データが無ければおかしいので終了
        if loaded_ms_data == {}:
            LM.loginfo("【エラー】原稿データがありません。")
            driver.close()
            driver.quit()
            #エラー：原稿データなし
            return Errors.MISSINGMSDATA.name
        else:
            # 原稿IDループ
            for ms_id in loaded_ms_data:
                #残ってるxlsmファイルを削除
                removexlsm = removeFilesFromDownloadFolder(SYSDOWNLOADDIRPATH)
                if(removexlsm):
                    LM.loginfo("DLフォルダから残存ファイルを削除しました。")
                elif(removexlsm == False):
                    pass
                # エラーチェック
                #if(len(loaded_ms_data[ms_id]["error"]) > 0):
                    #LM.loginfo("前回DLエラーあり")
                    # エラー復旧処理
                    #downloadMsData(driver, ms_id, loaded_ms_data, download)
                #else:
                    #LM.loginfo("前回エラーなし")
                #原稿が更新されたか:確認しないでグループ全巡回
                #if(isMsUpdate(driver, ms_id)):
                LM.loginfo(ms_id + "更新チェック開始")
                # 変数取得
                ms_title = re.sub(r'[\\/:*?"<>|]+', '／',
                                    loaded_ms_data[ms_id]["name"])  # 原稿タイトル
                ms_group_tabs = ""  # グループタブ
                # 原稿ページへ
                ms_page_url = PROMOTIONSYS_MSURL + "?mode=flier-select&id=" + ms_id  # idの原稿ページURL
                driver.get(ms_page_url)
                # DOM要素読み込み待ち
                WebDriverWait(driver, 20).until(
                    EC.presence_of_all_elements_located)
                #準備中・・・が消えてから次へ
                WebDriverWait(driver, 20).until(
                    EC.invisibility_of_element_located((By.CLASS_NAME, 'bootbox')))
                # タブ巡回
                try:
                    # タブ巡回:タブ要素があれば
                    if(len(driver.find_elements(By.CLASS_NAME, 'group-tab')) > 0):
                        ms_group_tabs = driver.find_elements(
                            By.CLASS_NAME, 'group-tab')
                        for i in range(len(ms_group_tabs)):
                            #残ってるxlsmファイルを削除
                            removexlsm = removeFilesFromDownloadFolder(SYSDOWNLOADDIRPATH)
                            if(removexlsm):
                                LM.loginfo("DLフォルダから残存ファイルを削除しました。")
                            elif(removexlsm == False):
                                pass
                            g_tab_element = ms_group_tabs[i].find_element(By.TAG_NAME, 'a')
                            ms_group_name = g_tab_element.text
                            ms_group_id = g_tab_element.get_attribute("data-key")
                            # まとめは無視
                            if not ms_group_id == "all":
                                # グループの初期化：グループのデータがない場合は追加。
                                if not ms_group_id in loaded_ms_data[ms_id]["group"]:
                                    LM.loginfo(ms_group_id + "を追加")
                                    loaded_ms_data[ms_id]["group"][ms_group_id] = {
                                        "name": ms_group_name, "update_time": "", "author": "",}
                                    # msdata.jsonに書き出し
                                    INSTANCE_OF_PANDAJSON.setDumpMsjson(loaded_ms_data)
                                # 現在の更新日取得：直前で新規作成の場合は空
                                ms_group_update_time = loaded_ms_data[ms_id]["group"][ms_group_id]["update_time"]
                                LM.loginfo("今{" + ms_group_id + ":" +
                                            ms_group_name + "}見てる")
                                # タブ遷移
                                try:
                                    # タブが1つの時はアクティブになってるのでクリックしない
                                    if(len(ms_group_tabs) > 1):
                                        clickWithJavascript(driver, driver.find_element(
                                            By.LINK_TEXT, g_tab_element.text))
                                        # グループの原稿テープルが表示されるまで待つ。id=table-g121（グループID）
                                        WebDriverWait(driver, 30).until(
                                            EC.visibility_of_element_located((By.ID, 'table-' + ms_group_id)))
                                except Exception as e:
                                    LM.loginfo(str(e))
                                    LM.loginfo('グループページエラー:次へ')
                                    continue
                                # グループが更新されたか
                                isupdate = IsUpdateMsWithOpenHistoryAll(driver, ms_group_id, ms_group_update_time)
                                # 更新なければ次のタブへ
                                if(isupdate == False):
                                    continue
                                else:
                                    # 原稿ダウンロード
                                    # 更新されてなければ次のグループへ
                                    # 履歴チェック
                                    ms_group_update_time = isupdate["update_time"]  # 更新日時上書き
                                    ms_group_update_author = isupdate["author"] # 更新者
                                    # 原稿ダウンロード
                                    if download_enable:
                                        for try_count in range(3): # 最大3回ダウンロードをトライ
                                            try:
                                                clickWithJavascript(driver, driver.find_element(By.LINK_TEXT, PROMOTIONSYS_DLTEXT))
                                                # DL開始待ち（準備中が消えるまで）ダウンロード開始まで時間かかるやつある。30秒、60秒、90秒と増やす。
                                                WebDriverWait(driver, 30 * try_count).until(EC.invisibility_of_element_located((By.CLASS_NAME, 'bootbox')))
                                            except Exception as e:
                                                LM.loginfo(str(e))
                                                LM.loginfo('【DL開始に失敗】' + "{" + try_count + "}" + "{" + ms_title + "}" + "{" + ms_group_name + "}")
                                            else:
                                                #成功なら抜ける
                                                break
                                        else:
                                            #全部失敗したら更新せず次へ
                                            LM.loginfo('【最大回数DLに失敗】' + "{" + ms_title + "}" + "{" + ms_group_name + "}")
                                            continue
                                        try:
                                            wait_file_download(SYSDOWNLOADDIRPATH, 60)  # クロームのDL待ち。60秒でエラー
                                            # DLファイル取得。存在するまで1秒毎に10回トライ。最新の.xlsmファイルをDLファイルとする。
                                            # このタイミングでの最新.xlsmがが選択されてしまう。
                                            # ここまでで、DL開始は保証され、Chromeの一時ファイルがなくなった事を確認しているが、外部からファイルを入れられるとまずい。
                                            # 現状DLフォルダを使用者に触れない様に隠蔽するしか方法がない。
                                            fname_before = False
                                            looplim = 0
                                            while(((fname_before == False) or (fname_before == None)) and looplim < 10):
                                                fname_before = getLatestDownloadedFileName(SYSDOWNLOADDIRPATH)
                                                time.sleep(1)
                                                looplim = looplim + 1
                                            if(looplim >= 10):
                                                LM.loginfo('ダウンロードファイルが正常に取得できてないよ')
                                            looplim = 0
                                            
                                            f_time_str_day = ms_group_update_time[5:10]
                                            f_time_str_time = ms_group_update_time[11:16]
                                            fname = re.sub(r'[\\/:*?"<>|-]+', '', f_time_str_day) + re.sub(r'[\\/:*?"<>|]+', '／', ms_title) + "【" + re.sub(
                                                r'[\\/:*?"<>|]+', '-', ms_group_name) + "】" + re.sub(r'[\\/:*?"<>|]+', '', f_time_str_time) + ".xlsm"
                                            #fname_after = os.path.join(DOWNLOADDIRPATH, fname)
                                            # fname_beforeがDL完了するまで待つ。
                                            wait_file_exist(fname_before, 30)

                                            # ファイル名変更して移動、上書き
                                            # カテゴリを判定して表示
                                            ms_category = INSTANCE_OF_CATEGORYJSON.what_category(ms_group_id)
                                            if not ms_category == False:
                                                LM.loginfo(ms_category)
                                            else:
                                                LM.loginfo("カテゴリ不明")

                                            #DLフォルダループ-START
                                            for dlpath in DOWNLOADDIRPATH:
                                                ms_folder_path = os.path.join(dlpath, ms_title)  # タイトル毎のフォルダ
                                                fname_after = os.path.join(ms_folder_path, fname)
                                                if os.path.isfile(fname_before):
                                                    if isConnected(dlpath):
                                                        # 原稿フォルダが存在しなければ作成。
                                                        my_makedirs(ms_folder_path)
                                                        #サーバー移動はshutil.move
                                                        shutil.copy(fname_before, fname_after)
                                                        LM.loginfo("【ファイル移動完了】" + fname_before)
                                                    else:
                                                        LM.loginfo("【panda:原稿移動】"+ dlpath + "に接続されていない")
                                                        driver.close()
                                                        driver.quit()
                                                        #エラー：DLフォルダ未接続
                                                        return Errors.DLFOLDERCONNECTION.name
                                                else:
                                                    LM.loginfo("【DLしたファイルが見つからない】" + fname_before)
                                            #DLフォルダループ-END

                                            # DL完了後処理
                                            # 原稿のupdatetimeを更新(あてにならない)
                                            loaded_ms_data[ms_id]["update_time"] = global_ms_json[ms_id]["update_time"]
                                            # グループのupdatetimeを更新
                                            loaded_ms_data[ms_id]["group"][ms_group_id]["update_time"] = ms_group_update_time
                                            # グループのauthorを更新
                                            loaded_ms_data[ms_id]["group"][ms_group_id]["author"] = ms_group_update_author
                                            # msdata.jsonに書き出し
                                            INSTANCE_OF_PANDAJSON.setDumpMsjson(loaded_ms_data)
                                            # history.xlsxに書き出し
                                            global_ed.updateHistory(ms_id, ms_group_name, ms_group_update_time, ms_group_update_author)
                                            # INSTANCE_OF_PANDAJSON.dumpsJsonData()
                                            # 更新分を格納
                                            addUpdateMsIDs(ms_id, ms_group_id)
                                        except Exception as e:
                                            LM.loginfo(str(e))
                                            LM.loginfo('DLエラー')
                                            continue
                            else:
                                LM.loginfo("まとめタブスルー")
                    else:
                        # 原稿更新されてるけど、グループタブないから、更新時間のみ入れとく。
                        LM.loginfo("グループタブがない")
                        # 原稿のupdatetimeを更新
                        loaded_ms_data[ms_id]["update_time"] = global_ms_json[ms_id]["update_time"]
                        # msdata.jsonに書き出し
                        INSTANCE_OF_PANDAJSON.setDumpMsjson(loaded_ms_data)

                except Exception as e:
                    #エラーをreturnしてmanagerで管理
                    LM.loginfo(str(e))
                    LM.loginfo('【深刻なエラー(GPtab）】システム管理者に報告してください。')
                    driver.close()
                    driver.quit()
                    #エラー：タブ巡回中エラー
                    return Errors.GROUPTAB.name
                    #sys.exit()
        
        LM.loginfo("【全部完了】終了します。")
        driver.close()
        driver.quit()
        #正常終了
        return Errors.NOTHING.name
    
    except Exception as e:
        #エラーをreturnしてmanagerで管理
        LM.loginfo(str(e))
        LM.loginfo("【深刻なエラー(all)】。システム管理者に報告してください。")
        driver.close()
        driver.quit()
        return Errors.MSTAB.name
        #sys.exit()
        #エラー：ログイン〜原稿ページループ

if __name__ == "__main__":
    panda()
