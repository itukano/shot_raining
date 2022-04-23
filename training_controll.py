'''
softalk
https://w.atwiki.jp/softalk/

ビープ音
https://sozorablog.com/beep/

0 winsound.Beep(262,1000)  #ド（262Hzの音を1000msec流す）
1 winsound.Beep(294,1000)  #レ（294Hzの音を1000msec流す）
2 winsound.Beep(330,1000)  #ミ（330Hzの音を1000msec流す）
3 winsound.Beep(349,1000)  #ファ（349Hzの音を1000msec流す）
4 winsound.Beep(392,1000)  #ソ（392Hzの音を1000msec流す）
5 winsound.Beep(440,1000)  #ラ（440Hzの音を1000msec流す）
6 winsound.Beep(494,1000)  #シ（494Hzの音を1000msec流す）
7 winsound.Beep(523,1000)  #ド（523Hzの音を1000msec流す）

python から外部プロセスの立ち上げ
https://boukenki.info/python-gaibu-program-process-application-kidou-jikkou-houhou/#Pythonrun

'''

import winsound
import subprocess
import time
import os

import numpy as np
import pandas as pd
import openpyxl as pyxl

#音関係の関数設定
eight_sound_list = []
for f in range(262, 530, 36):
    eight_sound_list.append(f) #アバウトなドレミファソラシド
#セット開始の音
def start_sound():
    for i in range(3):
        winsound.Beep(eight_sound_list[2], 400)
        time.sleep(0.8)

    winsound.Beep(eight_sound_list[5], 1500)
    time.sleep(0.3)
#もう少しでセット終了の音
def almost_sound():
    for i in range(3):
        for o in range(3):
            winsound.Beep(eight_sound_list[5], 100)
        time.sleep(0.3)
#セット終了の音
def goal_sound():
    winsound.Beep(eight_sound_list[5], 2000)

    for o in range(2):
        winsound.Beep(eight_sound_list[7], 100)



#運動の設定を開く
xls_path = "settings.xlsx"
if not os.path.exists(xls_path):
    print("can't find the path: ", xls_path)

try:
    if not os.access(xls_path, os.W_OK):
        print("prohibit to access file.")
        print("try to get permission.")
        os.chmod(xls_path, 755)

    print("try to open a excel file.")
    xls_table = pd.read_excel(xls_path)
except:
    print("can't get permittion.")
    print("give up to read the file: ", xls_path)
    sys.exit()

#運動の設定
subprocess.run([r"softalk\SofTalk.exe",  #念のため一回 softalk を立ち上げ
              "/X:1"])              
interval_elem = xls_table.query('name=="interval"')
interval_seconds = 0
if len(interval_elem) > 0:
    interval_elem = interval_elem.iloc[0]
    interval_seconds = interval_elem["seconds"]

almost_th = 0.8 #全体の何％が終わったところであと少しお知らせを行うか。鳴らしたくなければ 1 より大きくしておけばよい

#読み上げる文章を作る
def create_elem_text(elem):
    return elem["text"] + "。" + str(elem["seconds"]) + "びょう"

#テキストを読み上げる
def read_text(text):
    print(text)
    subprocess.run([r"softalk\SofTalk.exe", 
              "/W:" + text])
    t = len(text) / 4.0 #テキストの長さに合わせて適度に待つ。マジックナンバーは調整値
    time.sleep(t)

#筋トレの一項目の処理
def one_item(elem):
    set_time = elem["seconds"]
    als_flg = True #「もう少し」サウンドを鳴らしてよいかフラグ 
    for t in range(set_time):
        if t == 0:
            read_text(create_elem_text(elem))
            start_sound()
        elif (True == als_flg) and (t > set_time*almost_th):
            almost_sound()
            als_flg = False
        time.sleep(1)
        print(t, "/", set_time) 
     
    goal_sound()

#インターバル
def interval_time():
    if len(interval_elem) > 0:
        read_text(create_elem_text(interval_elem))
        time.sleep(interval_elem["seconds"]) 

#運動開始
print("start procedure")
total_time = sum(xls_table["repeat"] * xls_table["seconds"]) + ((len(xls_table)-1)*interval_seconds)
read_text("うんどうじかんは、やく" + str(int(total_time/60)) + "ぷんです")
time.sleep(5)

repeat_num = max(xls_table["repeat"])
print(repeat_num)
for r in range(repeat_num):
    read_text("だい"+str(r+1)+"せっと。")
    for i in range(len(xls_table)):
        print("process: ", i+1, "/", len(xls_table))
        elem = xls_table.loc[i]
        if (elem["name"] != "interval") and (r < elem["repeat"]):
            print(elem["name"])
            one_item(elem)

            if (r == repeat_num-1) and (i != len(xls_table)-1):
                interval_time()

read_text("とれーにんぐ。しゅうりょうです。")


