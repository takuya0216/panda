from logging import NullHandler
import os
import sys
from collections import defaultdict
from attr import NOTHING
from dotenv import load_dotenv
from enum import Enum

# ダウンロードフォルダ（サーバー）に接続できているかチェック。
# 接続できてればTrue、接続できてない場合はそのフォルダパスを返す
def isConnectedDLPath():
    errorlist = []
    for dlpath in DLROOTPATH:
        if not (os.path.exists(dlpath)):
            errorlist.append(dlpath)
    
    if errorlist == []:
        return True
    else:
        return errorlist

#接続確認引数版
def isConnected(dlpath):
    if(os.path.exists(dlpath)):
        return True
    else:
        return False

# システム全体で使用する定数を管理する。書き換えられないよう注意
# UPDATEDMSIDSに追加する。
def addUpdateMsIDs(ms_id, group_id):
    UPDATEDMSIDS[ms_id].append(group_id)
    print("update_append:" + ms_id + ":" + group_id)

# UPDATEDMSIDSを初期化する。（空にする）
def clearUpdateMsIDS():
    UPDATEDMSIDS.clear()

def getDesktopPath():
    if os.name == 'nt':
        desktop_dir = os.path.join(os.path.join(
            os.environ['USERPROFILE']), 'Desktop')
    else:
        desktop_dir = os.path.join(os.path.join(os.environ['HOME']), 'Desktop')

    return desktop_dir

# ファイルパス関係
APP_PATH = (
    sys._MEIPASS
    if getattr(sys, "frozen", False)
    else os.path.dirname(os.path.abspath(__file__))
)

#設定データ(機密情報)を環境変数にロード（.env）
load_dotenv(os.path.join(APP_PATH, 'data/conf/.env'))

#DLフォルダのルート。サーバーのパスなど接続確認用。VolumesはMacのマウントフォルダ。
DLROOTPATH = os.environ['DLROOTPATH'].split(',') #,区切りのデータを配列化。
DLFOLDERNAME = os.environ['DLFOLDERNAME']

#### Value for panda_manager.py ####
# 更新されたms_idとgroupidの辞書。pandaで格納。
UPDATEDMSIDS = defaultdict(list)

#### Value For panda.py ####
DOWNLOADDIRPATH = []
for path in DLROOTPATH:
    DOWNLOADDIRPATH.append(os.path.join(path, DLFOLDERNAME)) #ダウンロードしたファイルの移動先
    print(path)

SYSDOWNLOADDIRPATH = os.path.join(APP_PATH, 'data/download/')  # ブラウザからのダウンロードフォルダのパス
PROMOTIONSYS_MSURL = os.environ['MS_URL']
PROMOTIONSYS_LOGINURL = os.environ['LOGIN_URL']
PROMOTIONSYS_ID = os.environ['PROMOTIONSYS_ID']
PROMOTIONSYS_PW = os.environ['PROMOTIONSYS_PW']
PROMOTIONSYS_DLTEXT = "グループファイルダウンロード"

#### Value for excel_manager.py ####
# 履歴エクセルのパス：excel_manager.py
WORKBOOKPATH = os.path.join(APP_PATH, "data/history.xlsx")

#### Value for panda_json.py ####
MSJSONPATH = os.path.join(APP_PATH, "data/sys/msdata.json")  # 原稿データ保存パス
MSCATEGORYJSONPATH = os.path.join(APP_PATH, "data/sys/category.json")  # 原稿データ保存パス

#### Value for logmanager.py ####
LOGDIRPATH = os.path.join(APP_PATH, 'data/logs/')  # ログフォルダのパス
# ログファイルのパス
LOGDATAPATH = os.path.join(APP_PATH, 'data/logs/{}.logs')
# ログ設定ファイルのパス
LOGCONFIGPATH = os.path.join(APP_PATH, 'data/conf/log_config.json')

# エラークラス（列挙型）
# エラー名：エラー表示テキスト
class Errors(Enum):
    DLFOLDERCONNECTION = "ダウンロードフォルダにアクセスできません。フォルダへの接続を確認して実行してください。"
    MISSINGSYSDLFOLDER = "システムのダウンロードフォルダが見つかりません。管理者に報告してください。"
    INITWEBDRIVER = "WEBドライバーの初期化に失敗しました。ネットワーク接続を確認してください。何度か実行しても改善しない場合は、管理者へ報告してください。"
    MISSINGMSDATA = "原稿データの取得に失敗。管理者に報告してください。"
    GROUPTAB = "グループタブ巡回に失敗。管理者に報告してください。"
    MSTAB = "原稿ページループに失敗。管理者に報告してください。"
    LOGIN = "ログインに失敗。管理者に報告してください。"
    NOTHING = "エラーなし"
