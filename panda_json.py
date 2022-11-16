import json
import os
from re import A
import sys
from panda_value import MSJSONPATH, MSCATEGORYJSONPATH

"""
原稿出力用のjsonデータを扱うクラス
現状コンストラクトとゲット、セットのみ。
インスタンス変数：jsondata=Jsonのデータ。コンストラクタで生成。
データ構成：原稿ID>グループIDの入れ子。それぞれ、名前と更新日を持つ。
原稿IDは、システムの表示範囲に増えたら追加、消えたら削除する。
グループは、増えたら追加、消えても削除しない。
エラーリストは、現在使用していない。
{
     "33448":{
        "name":"【第41週 12/10号 春日井店】生活応援",
        "update_time":"2021-11-15 16:28:32"
        "error":[],
        "group":{
            "all":{
                "name":"まとめ"
                "update_time":"2021-11-15 16:28:32"
                "author":"石田"
            }
            "g116":{
                "name":"雑貨",
                "update_time":"2021-11-15 16:28:32"
                "author":"石田"
            }
        }
    }
}
"""

class msDatas():

    #dataFilePath = os.path.join(os.path.dirname(__file__), "data/msdata.json") #原稿データ保存パス
    dataFilePath = MSJSONPATH #原稿データ保存パス

    jsondata = {} #jsonデータ

    #コンストラクタ：jsonデータ（辞書型）を渡す
    def __init__ ( self ) :
        print("コンストラクト：msDatas")
        self.init()

    #デストラクタ
    def __del__ ( self ) :
        print("デストラクト:msDatas")

    #初期化。msdata.jsonファイルから読み込み、データがなければ空で作成。jsondata（クラス変数）にjsonデータをセット
    def init(self):
        #原稿Jsonロード
        if(os.path.isfile(self.dataFilePath)):
            json_open = open(self.dataFilePath, 'r' , encoding='utf-8')
            try:
                self.jsondata = json.load(json_open)
            except Exception as e:
                print('ファイルがロードできません。書式を確認してください。')
                sys.exit()
        else:
            print(self.dataFilePath + "がありません")
            #データ新規作成
            with open(self.dataFilePath, "w", encoding='utf-8') as f:
                json.dump(self.jsondata, f, ensure_ascii=False, indent=4)
        
    #jsonファイルから読み込み、辞書型として返す
    def loadFromJson(fpath):
        json_open = open(fpath, 'r', encoding='utf-8')
        return(json.load(json_open))

    #msdata.jsonが存在するか
    def isExistJson(self):
        if(os.path.isfile(self.dataFilePath)):
            return True
        else:
            return False

    #jsonデータを取得
    def getJsonData(self):
        return self.jsondata

    #jsonデータをセット
    def setjsonData(self, JsonData):
        self.jsondata = JsonData

    #msdata.jsonファイルから読み込み、jsondata（インスタンス変数）にデータセット
    def setjsonDataFromJson(self):
        #原稿Jsonロード
        if(os.path.isfile(self.dataFilePath)):
            json_open = open(self.dataFilePath, 'r' , encoding='utf-8')
            self.jsondata = json.load(json_open)
        else:
            print("jsonファイルがありません")
            sys.exit()

    #jsonデータをコンソール出力
    def dumpsJsonData(self):
        print(json.dumps(self.jsondata,ensure_ascii=False, indent=4))

    #jsonデータを書き出す
    def dumpJsonData(self, fpath):
        with open(fpath, "w" , encoding='utf-8') as f:
            json.dump(self.jsondata,f,ensure_ascii=False, indent=4)

    #jsonデータをセットして、Msdata.jsonに書き出す
    def setDumpMsjson(self, JsonData):
        self.jsondata = JsonData
        with open(self.dataFilePath, "w", encoding='utf-8') as f:
            json.dump(self.jsondata, f, ensure_ascii=False, indent=4)

"""
原稿のカテゴリJSONクラス
JSONデータの操作はmsDataを継承
SSM・DC・HF・GMの振り分け
{
    "SSM":{
        "group":{
            "g116":{
                "name":"雑貨",
            }
        }
    }
    "DC":{
        "group":{
            "g116":{
                "name":"雑貨",
            }
        }
    }
    "HF":{
        "group":{
            "g116":{
                "name":"雑貨",
            }
        }
    }
    "GM":{
        "group":{
            "g116":{
                "name":"雑貨",
            }
        }
    }
}
"""
class msCategories(msDatas):

    def __init__(self):
        self.dataFilePath = MSCATEGORYJSONPATH #jsonデータのパスを上書き
        super().__init__() #jsondataがdataFilePathのデータで初期化される。

    #ここから固有のメソッド
    def print_filepath(self):
        print(self.dataFilePath)
    
    #groupidのカテゴリを返す。カテゴリに登録がない場合はfalse
    def what_category(self, groupid):
        for ms_category in self.jsondata:
            if groupid in self.jsondata[ms_category]["group"]:
                return ms_category
        else:
            return False

    #groupidとg_nameでcategoryを更新。登録がない場合のみ更新。
    def update_category_json(self, groupid, g_name):
        ms_category = self.what_category(groupid)
        if ms_category == False:
            #カテゴリが登録ないので追加
            print("カテゴリがわかりません。入力してください。:")
            this_category = input()
            if not this_category in self.jsondata:
                print("そんなカテゴリはねーよ。SSM、GM、HF、DCのどれかで入力しろ。")
                return
            else:
                print(this_category + "ねー、了解。")
                self.addGroupAndDump(this_category, groupid, g_name)
        else:
            print(ms_category + "に登録済みです。")
    
    #カテゴリにグループを加えて、category.jsonも更新
    def addGroupAndDump(self, ms_category, groupid, g_name):
        category_json = self.getJsonData()
        category_json[ms_category]["group"][groupid] = {"name":g_name}
        self.setDumpMsjson(category_json)

### Instance for all ####
# msDatasインスタンスコンストラクト
INSTANCE_OF_PANDAJSON = msDatas()
# msCategoriesインスタンスコンストラクト
INSTANCE_OF_CATEGORYJSON = msCategories()