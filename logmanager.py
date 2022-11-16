from logging import getLogger, config
import json
import datetime
import os
import time
from panda_value import LOGCONFIGPATH
from panda_value import LOGDATAPATH
from panda_value import LOGDIRPATH

# 実行ログをテキスト出力するクラス
# config設定のloggerインスタンスを作成。
# ログを撮りたいタイミングで、logMangerのクラスメソッドloginfo(str)でログ出力。
# LOGCONFIGPATHにコンフィグJsonファイルを置く必要あり。

class logManager():
    logger = getLogger(__name__)
    def __init__(self ,logger_name):
        print("コンストラクト:logManager")
        with open(LOGCONFIGPATH, 'r') as f:
            log_conf = json.load(f)

        # ファイル名をタイムスタンプで作成
        log_conf["handlers"]["fileHandler"]["filename"] = \
            LOGDATAPATH.format(datetime.datetime.now(datetime.timezone.utc).strftime("%Y%m%d%H%M%S"))
        # 設定ロード
        config.dictConfig(log_conf)
        # コンストラクト
        logManager.logger = getLogger(logger_name)
        # 古いログを削除
        logManager.removeLogs()

    # デストラクタ
    def __del__ ( self ) :
        print("デストラクト:logManager")
    
    # INFOレベルでログファイルにstrをログ出力
    def loginfo(str):
        logManager.logger.info(str)
    
    # 指定期日より前のログを削除する。
    # データ保持期間を7日間に設定。
    def removeLogs():   
        workdir = LOGDIRPATH
        now = time.time()
        old = now - 7 * 24 * 60 * 60
        for f in os.listdir(workdir):
            path = os.path.join(workdir, f)
            #ファイルが存在して、.logsファイルなら削除
            if os.path.isfile(path) and f.endswith('.logs'):
                stat = os.stat(path)
                if stat.st_ctime < old:
                    logManager.loginfo("removing: " + path)
                    os.remove(path) # uncomment when you will sure :)