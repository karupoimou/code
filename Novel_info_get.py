#なろう小説情報取得
import re
import os
import requests
import time
import xlrd
import openpyxl as opx
from bs4 import BeautifulSoup
import datetime
import pandas as pd


############　ここだけ変更すればOK　######################

#ここに作品のNcodeを指定する
ncode="n2919fm"

#なろう本家なら「１」、ノクターンノベルズなら「0」に設定する
isNarou=1

##########################################################

#編集・追記していくファイル
path = 'check_novel_info.xlsx'
sheet_name=ncode

#記録時間を指定
now = datetime.datetime.now()
dt_now = datetime.datetime.now()
nowtime = dt_now.strftime('%Y-%m-%d %H:%M:%S')
filename = '{0:%Y%m%d}_{0:%H%M}'.format(now)

#ユーザーエージェントの設定（設定必須）
headers = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64; rv:61.0) Gecko/20100101 Firefox/61.0"}
cookie = {'over18': 'yes'}  # Xサイト用のクッキー

#データフレーム用
all_list=[]
columns_name=["取得日時","感想","レビュー","ブックマーク登録","総評価","ポイント評価(文章)","ポイント評価(ストーリー)"]

def load_excel():

    #ファイルがない場合は生成
    check_excel_file()

    # ncode読み込みファイル指定,ここでncodeが入ったファイルを指定する
    wb = xlrd.open_workbook(path)

    #追記するファイルを読み込む読み込む
    sheets = wb.sheets()
    sheet = wb.sheet_by_name(sheet_name)
    prev_set = sheet.col_values(0)

    for i in range(len(prev_set)-1):
        i=i+1
        temp_list=[]
        temp_list.append(sheet.cell_value(i,1))
        temp_list.append(sheet.cell_value(i,2))
        temp_list.append(sheet.cell_value(i,3))
        temp_list.append(sheet.cell_value(i,4))
        temp_list.append(sheet.cell_value(i,5))
        temp_list.append(sheet.cell_value(i,6))
        temp_list.append(sheet.cell_value(i,7))

        all_list.append(temp_list)

def check_excel_file():
    if os.path.exists(path)==False:
        df = pd.DataFrame(all_list)
        df.to_excel(path, sheet_name=sheet_name)
    else:
        pass

def set_url():
    if isNarou==1:
        url ="https://ncode.syosetu.com/novelview/infotop/ncode/%s/" %ncode
    else:
        url = 'https://novel18.syosetu.com/novelview/infotop/ncode/%s/' %ncode
    return url

def get_novel_info():

    url=set_url()

    response = requests.get(url=url, headers=headers, cookies=cookie)
    html = response.content
    soup = BeautifulSoup(html, "lxml")

    temp_list=[]

    #データ取得日時
    temp_list.append(filename)

    #感想数抽出
    sp=soup.select('table td')[-10].text
    sp = re.sub("\\D", "", sp)
    temp_list.append(int(sp))

    #レヴュー数
    sp=soup.select('table td')[-9].text
    sp = re.sub("\\D", "", sp)
    temp_list.append(int(sp))

    #ブックマーク数
    sp=soup.select('table td')[-8].text
    sp = re.sub("\\D", "", sp)
    temp_list.append(int(sp))

    #総評価ポイント
    sp=soup.select('table td')[-7].text
    sp = re.sub("\\D", "", sp)
    temp_list.append(int(sp))

    #文章ポイント
    sp=soup.select('table td')[-6].text
    index1=sp.find('p')
    pt_bunshou=sp[0:index1]#文章評価

    index2=sp.find('：')+1
    index3=sp.find('p',index1)
    pt_story=sp[index2:index2+index3]

    #ここでカンマを外す処理
    pt_bunshou = pt_bunshou.replace(",", "")
    pt_story = pt_story.replace(",", "")

    temp_list.append(int(pt_bunshou))
    temp_list.append(int(pt_story))

    all_list.append(temp_list)

#関数の実行
load_excel()
get_novel_info()

#############以下エクセルシートに書き込む処理################

df = pd.DataFrame(all_list,columns=columns_name)#pandasのデータフレームに収納

# 対象ファイルのExcelWriterの呼び出し
EXL=pd.ExcelWriter(path, engine='openpyxl')

# ExcelWriterに既存の対象ファイルを読み込ませる
EXL.book=opx.load_workbook(path)
# 既存のsheet情報を読み込ませる
EXL.sheets=dict((ws.title, ws) for ws in EXL.book.worksheets)

# ExcelWriterを用いて新規シートにDataFrameを保存 ここで新規分を書き込み
df.to_excel(EXL, sheet_name=sheet_name)

# ExcelWriterの処理を保存
EXL.save()
