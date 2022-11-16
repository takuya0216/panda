import datetime
from re import T
import schedule
import threading
import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter.ttk import *
from tkinter import messagebox
import time
from panda import panda
from logmanager import logManager as LM
from panda_value import *
from panda_json import *

# datetimeを文字列フォマットして返す。時間：分：秒
def datetimeTostring(dt):
    return dt.strftime('%H時%M分%S秒')

class JobManager:

    MONITOR_PERIOD = 60  # 実行間隔：秒
    newInterval = MONITOR_PERIOD  # インターバル更新用

    def __init__(self, root):

        #実行可能か確認
        #サーバーに接続チェック
        isConnected = isConnectedDLPath()
        if(isConnected == True):
            LM.loginfo("【manager】DLフォルダ接続OK!")
        else:
            errorstr = ""
            for errorPath in isConnected:
                LM.loginfo( "【manager】" + errorPath + "にアクセスできません。")
                errorstr = errorstr + "【" + errorPath + "】"
            messagebox.showerror("エラー", errorstr + "にアクセスできません。フォルダへの接続を確認して実行してください。")
            sys.exit()

        self.root = root
        self.running = True
        self.job_running = False
        self.job_enable = False
        self.error_state = Errors.NOTHING
        self.state = tk.StringVar()   

        # GUIの準備
        self.root.title(u"イオン原稿出力")
        self.root.geometry("429x199")
        self.state.set("システム停止中")

        frame = tk.Frame(root)
        frame.pack()
        button_start = ttk.Button(frame, text="Start", command=self.start)
        button_start.pack(side=tk.LEFT)
        button_stop = ttk.Button(frame, text="Stop", command=self.stop)
        button_stop.pack(side=tk.LEFT)
        button = ttk.Button(frame, text="Quit", command=self.quit)
        button.pack(side=tk.LEFT)
        label_state = tk.Label(root, textvariable=self.state)
        label_state.pack(anchor='center', expand=1)
        label_interbal = tk.Label(frame, text="実行間隔(秒):10秒以上")
        label_interbal.pack()
        self.entry = tk.Entry(frame)
        self.entry.insert(tk.END, self.MONITOR_PERIOD)
        self.entry.pack()
        # %s は変更前文字列, %P は変更後文字列を引数で渡す
        vcmd = (self.entry.register(self.validation), '%P')
        # Validationコマンドを設定（'key'は文字が入力される毎にイベント発火）
        self.entry.configure(validate='key', vcmd=vcmd)
        button_accept = ttk.Button(frame, text="適用", command=self.setInterval)
        button_accept.pack()

        # ジョブのスケジューリング
        schedule.every(self.MONITOR_PERIOD).seconds.do(self.job).tag("panda")
        # 次回の実行時間を更新
        self.setNextTime()

        # スレッドの開始
        self.thread = threading.Thread(target=self.run_monitor)
        self.thread.start()

        # macのテキストが表示されないバグフィックス：ウィンドウサイズを+1する
        self.root.update()
        self.root.after(0, self.fix)

    def __del__(self):
        print("デストラクト:pandaManager")

    def run_monitor(self):
        while self.running:
            schedule.run_pending()
            if self.job_enable:
                self.state.set("【システム稼働中】\n" + str(self.MONITOR_PERIOD) + "秒ごとに実行されます。\n" + "次は、【" + self.nextTime + "】頃に実行されます。")
            else:
                self.state.set("システム停止中")
            time.sleep(0.2)
        self.root.quit()

    # 定期実行ジョブ
    def job(self):
        if self.job_enable:
            isConnected = isConnectedDLPath()
            if(isConnected == True):
                LM.loginfo("【manager】DLフォルダ接続OK!")
                self.job_running = True
                self.state.set("【ジョブを実行中】\n" + "次は、【" + self.nextTime + "】頃に実行されます。")
                
                #ジョブ実行：エラーステート更新
                self.error_state = panda()
                
                #pandaからのエラー処理
                self.manageError()

                #後処理
                self.afterJob()
                
            else:
                errorstr = ""
                for errorPath in isConnected:
                    LM.loginfo( "【manager】" + errorPath + "にアクセスできません。")
                    errorstr = errorstr + "【" + errorPath + "】"
                messagebox.showerror("エラー", errorstr + "にアクセスできません。フォルダへの接続を確認して実行してください。")
                self.job_enable = False
        
        self.job_running = False
    

    #エラー処理
    def manageError(self):
        #エラーなし
        if  self.error_state == Errors.NOTHING.name:
            pass
        #DLフォルダ接続エラー
        elif self.error_state == Errors.DLFOLDERCONNECTION.name:
            messagebox.showerror("エラー", Errors.DLFOLDERCONNECTION.value)
            #システム停止へ
            self.job_enable = False
        #ドライバー初期化エラー
        elif self.error_state == Errors.INITWEBDRIVER.name:
            messagebox.showerror("エラー", Errors.INITWEBDRIVER.value)
            #エラー表示のみ
        #msdata.jsonエラー
        elif self.error_state == Errors.MISSINGMSDATA.name:
            messagebox.showerror("エラー", Errors.MISSINGMSDATA.value)
            #システム停止へ
            self.job_enable = False
        #グループタブループエラー
        elif self.error_state == Errors.GROUPTAB.name:
            messagebox.showerror("エラー", Errors.GROUPTAB.value)
            #システム停止へ
            self.job_enable = False
        #原稿ループエラー
        elif self.error_state == Errors.MSTAB.name:
            messagebox.showerror("エラー", Errors.MSTAB.value)
            #システム停止へ
            self.job_enable = False
        #ログインエラー
        elif self.error_state == Errors.LOGIN.name:
            messagebox.showerror("エラー", Errors.LOGIN.value)
            #システム停止へ
            self.job_enable = False
        #システムDLフォルダエラー
        elif self.error_state == Errors.MISSINGSYSDLFOLDER.name:
            messagebox.showerror("エラー", Errors.MISSINGSYSDLFOLDER.value)
            #システム停止へ
            self.job_enable = False
        #設定忘れでもステートは表示されるように
        else:
            messagebox.showerror("エラー", self.error_state)
            #システム停止へ
            self.job_enable = False


    #ジョブの後処理
    def afterJob(self):
        #原稿更新表示
        if(len(UPDATEDMSIDS) > 0):
            INSTANCE_OF_PANDAJSON.setjsonDataFromJson() #msdata.jsonからjsonデータを更新
            current_msdata = INSTANCE_OF_PANDAJSON.getJsonData() #jsonデータを取得
            LM.loginfo("【更新あり】\n")
            for ms_id in UPDATEDMSIDS:
                LM.loginfo(current_msdata[ms_id]["name"])
                for ms_group_id in UPDATEDMSIDS[ms_id]:
                    LM.loginfo(current_msdata[ms_id]["group"][ms_group_id]["name"] + ":" + current_msdata[ms_id]["group"][ms_group_id]["update_time"] + "\n")
        else:
            LM.loginfo("更新はありませんでした。")
        
        # スケジュールの更新
        if not self.MONITOR_PERIOD == self.newInterval:
            self.updateSchedule()
        self.setNextTime()
    
    def start(self):
        #ジョブ実行中はスタートできない。
        if self.job_running == True:
            messagebox.showinfo("お知らせ","ジョブが実行中です、完了するまでお待ちください。")
        else:
            isConnected = isConnectedDLPath()
            if(isConnected == True):
                if self.job_enable == False:
                    #スケジュール更新
                    if not self.MONITOR_PERIOD == self.newInterval:
                        self.updateSchedule()
                    self.setNextTime()
                    LM.loginfo("システムを稼働")
                    self.job_enable = True
            else:
                errorstr = ""
                for errorPath in isConnected:
                    LM.loginfo( "【manager】" + errorPath + "にアクセスできません。")
                    errorstr = errorstr + "【" + errorPath + "】"
                messagebox.showerror("エラー", errorstr + "にアクセスできません。フォルダへの接続を確認して実行してください。")
                self.job_enable = False

    def stop(self):
        if(self.job_enable):
            self.state.set("システム停止中")
            LM.loginfo("システム停止")  
        self.job_enable = False
        
    def quit(self):
        if(self.job_running == True):
            messagebox.showinfo("お知らせ","ジョブが実行中です、完了するまでお待ちください。")
        else:
            answear = messagebox.askquestion('アプリケーションの終了','アプリケーションを終了してもよろしいですか', icon='warning')
            if answear == 'yes':
                self.running = False
            else:
                messagebox.showinfo('お知らせ','アプリケーション画面に戻ります')
            

    # 次の実行時刻をセットする。（文字列）
    def setNextTime(self):
        self.nextTime = datetime.datetime.now(
        ) + datetime.timedelta(seconds=self.MONITOR_PERIOD)
        self.nextTime = datetimeTostring(self.nextTime)

    # 文字列検証関数:入力後の文字列が、10進数の1文字以上の数字かどうか
    def validation(self, after_word):
        return ((after_word.isdecimal()))

    # インターバル時間の値を変更
    def setInterval(self):
        # 入力されたインターバルでスケジュールし直し
        self.newInterval = int(self.entry.get())
        if not self.newInterval >= 10:
            messagebox.showinfo("エラー", "実行間隔は10以上で設定してください。")
            self.newInterval = self.MONITOR_PERIOD
        else:
            messagebox.showinfo("お知らせ", str(
                self.newInterval) + "秒が設定されました。次回から設定されます。")

    # 実行中スケジュールを削除して、新しいスケジュールを作成
    def updateSchedule(self):
        schedule.clear("panda")
        LM.loginfo("スケジュール削除")
        # ジョブのスケジューリング
        schedule.every(self.newInterval).seconds.do(self.job).tag("panda")
        LM.loginfo("スケジュール再設定されました")
        self.MONITOR_PERIOD = self.newInterval  # update

    # macのバグフィックス
    def fix(self):
        a = self.root.winfo_geometry().split('+')[0]
        b = a.split('x')
        w = int(b[0])
        h = int(b[1])
        self.root.geometry('%dx%d' % (w+1, h+1))

def main():
    LM(__name__)
    LM.loginfo("【ログ出力開始】")
    root = Tk()
    JobManager(root)
    root.mainloop()


if __name__ == "__main__":
    main()
