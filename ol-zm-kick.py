""" outlook の今日の予定から URL を取り出して、zoom か Teams を起動する ver1.2  """

# from nturl2path import url2pathname
import webbrowser
import win32com.client
import datetime
import time
import re
import subprocess

def str2time(str_time): # %H:%M:%S 形式の文字列をdatetimeオブジェクトに変換する
    return datetime.datetime.strptime(str_time, "%H:%M:%S")

tm_URL = "<https:/[\w/:%#\$&\?\(\)~\.=\+\-]+teams.microsoft.com[\w/:%#\$&\?\(\)~\.=\+\-]+>"
kick_url = ''
timer = 0

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

calendar = outlook.GetDefaultFolder(9).items
calendar.Sort("[Start]")
calendar.IncludeRecurrences = "True"

select_items = [] # 指定した期間内の予定を入れるリスト

# 予定を抜き出したい期間を指定
today_date = datetime.date.today() # 今日だけ
now_time = datetime.datetime.now() # 今の時間
'''
today_date = datetime.date.fromisoformat('2022-06-20') # test 
now_time = datetime.datetime.fromisoformat('2022-06-17 18:00:00') # test
#'''
scan_start_time = now_time + datetime.timedelta(hours=2) # 今の時間 + 2H で開始する予定
scan_start = scan_start_time.strftime('%H:%M:%S') #
scan_end_time = now_time + datetime.timedelta(minutes=15) # 今の時間 + 15min 
scan_end = scan_end_time.strftime('%H:%M:%S') #15min の加算補正あり

# restrict appointments to specified range
calendar = calendar.Restrict("[Start] >= '" + str(today_date) +
                             "' AND [END] <= '" + str(today_date + datetime.timedelta(days=1)) + "'")

# 今日のデータを取り出し
for item in calendar:
    if today_date == item.start.date():
        select_items.append(item)
    if today_date < item.start.date():
        break 

# 抜き出した予定の詳細を表示
for select_item in select_items:
    if ( select_item.body == '' or select_item.body == ' ' or 
         select_item.subject.startswith('Canceled:') or
         select_item.subject.startswith('キャンセル済み:') ) : 
        print("対象外会議あり：", select_item.subject)
        continue

    if  select_item.ResponseStatus != 3 and select_item.ResponseStatus != 1 : # 3=承諾済み 1=自分の予定
        print("未承認会議あり：", select_item.subject)
        continue
     #本文が承諾済みでは無いので、次の予定に飛ぶ
     # https://docs.microsoft.com/ja-jp/office/vba/api/outlook.olresponsestatus
    
    start_time = select_item.start.time().strftime('%H:%M:%S')
    end_time = select_item.end.time().strftime('%H:%M:%S')

    if scan_start >= start_time and scan_end < end_time :  #開始2H前～終了15分前まで
        print("該当会議あり：", select_item.subject)
        lines = select_item.body.split()
    else :
        continue
    # 次の予定へ

    print (select_item.MeetingWorkspaceURL)

    # URL を探す
    zm_key_idx = -1
    zoom_id = ''
    zoom_pc = ''
    
    for zm_key_idx, line in enumerate(lines) :
        if re.match(tm_URL, line):
            kick_url = line[1:len(line)-1]
            break
        if line == 'ミーティングID:' : 
            zoom_id = lines[zm_key_idx+1] + lines[zm_key_idx+2] + lines[zm_key_idx+3]
            if lines[zm_key_idx+4] == 'パスコード:' or lines[zm_key_idx+4] == 'パスワード:' :
                zoom_pc = lines[zm_key_idx+5]
        if line == 'ID):' : 
            zoom_id = lines[zm_key_idx+1] + lines[zm_key_idx+2] + lines[zm_key_idx+3]
            if lines[zm_key_idx+4] == 'パスコード(Passcode):' or lines[zm_key_idx+4] == 'パスワード(Password):' :
                zoom_pc = lines[zm_key_idx+5]
        if zoom_id != '' : #id が取り出せた
            kick_url = 'zoommtg://zoom.us/join?confno=' + zoom_id + '&pwd=' + zoom_pc
            break
    break #kick_url を見つけたら、先の予定は見ない   

# 結果の判定
if kick_url == "" : #それらしい URL が無い
    print ('近い時間の web会議がありません')
    timer = input ('何分のタイマーをかけますか？')
    if ( timer != '' and timer != '0' ) :
        if timer.isdecimal() :
            subprocess.Popen(r'C:\Program Files (x86)\Hourglass\hourglass.EXE ' + timer + 'm')
    exit()

now_time = now_time.strftime('%H:%M:%S')  #実時間
#now_time = '09:56:00' # debug 用

sleep_time = int((str2time(start_time) - str2time(now_time)).total_seconds() - 40) #分と秒で 40秒前に起動
if sleep_time > 0 : #残りが40秒以上の場合、タイマをかける
    print (f'起動まで { int(sleep_time / 60) } 分 待ちます' )
    subprocess.Popen(r'C:\Program Files (x86)\Hourglass\hourglass.EXE -e on -r off --title <' + select_item.subject.replace(' ', '') + '> ' +str(sleep_time) + 's')
    # https://chris.dziemborowicz.com/apps/hourglass/#command-line-arguments
    time.sleep(sleep_time)
 
webbrowser.open(kick_url)