'mac command' 
'''
版本替換要換:
def runcmd:
    if 'ddm' in cmd[0:3] or 'DDM' in cmd[0:3]:
    if './DDM' in cmd:
        
version = macos
command typing , runcmd need to add ddpm to ddm


'''
#測試無網路狀態
import ast
import sys
import json
import keyboard               
import pyautogui
import time   
import re
import socket
import pandas as pd
from datetime import datetime as dt2 
# import qdarkstyle
import os
import csv
from pathlib import Path
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QIcon, QPixmap
import gspread
from google.oauth2.service_account import Credentials
import requests
from pyrebase import pyrebase
from datetime import datetime as dt2
from sys import platform
import webbrowser
import pygsheets
python_path = os.getcwd()
beta_name = python_path+'/res'+"/beta2.csv"
if platform == "linux" or platform == "linux2":
    version = 'linux'
elif platform == "darwin":
    version = 'macos'
elif platform == "win32":
    version = 'window'
print('version',version)
df = pd.read_csv(python_path+'/res/setup_command.csv')
if version == 'macos':
    cmd_cd = df.mac_command[0]
    cmd_mkdir ='mkdir temp'
    cmd_logcreate =df.mac_command[1]
    cmd_logfile = df.mac_command[2]
    cmd_dialogfile = df.mac_command[3]
    cmd_getAttritube = df.mac_command[4]
    terminal_cmd = df.mac_command[5]
    input_filter = df.mac_command[6]
    preset_filter = df.mac_command[7]
    single_input_read = df.mac_command[8]
    multi_input_read = df.mac_command[9]
    single_preset_read = df.mac_command[10]
    multi_preset_read = df.mac_command[11]
    multiple_task = int(df.mac_command[12])
    reset_language = df.mac_command[13]
    pip_off = df.mac_command[14]
    reset_input_2 = df.mac_command[15]
    reset_input_1 = df.mac_command[16]
    
elif  version == 'window':
    cmd_cd = df.window_command[0]
    cmd_mkdir ='mkdir C:\\temp'
    cmd_logcreate = df.window_command[1]
    cmd_logfile = df.window_command[2]
    cmd_dialogfile = df.window_command[3]
    cmd_getAttritube = df.window_command[4]
    terminal_cmd = df.window_command[5]
    input_filter = df.window_command[6]
    preset_filter = df.window_command[7]
    single_input_read = df.window_command[8]
    multi_input_read = df.window_command[9]
    single_preset_read = df.window_command[10]
    multi_preset_read = df.window_command[11]
    multiple_task = int(df.window_command[12])
    reset_language = df.window_command[13]
    pip_off = df.window_command[14]
    reset_input_2 = df.window_command[15]
    reset_input_1 = df.window_command[16]

subinputname = ast.literal_eval(df.mac_command[17])
window = ast.literal_eval(df.mac_command[18])
presetname1 = ast.literal_eval(df.mac_command[19])
preset1 = ast.literal_eval(df.mac_command[20])
presetname3 = ast.literal_eval(df.mac_command[21])
preset3 = ast.literal_eval(df.mac_command[22])


def ddpm_ddm(target,version):
    '''
    MacOS , Window 的前綴變換
    '''
    target = target.replace('DDM',version)
    target = target.replace('DDPM',version)
    print('target',target)
    return target
def clean(string):
    '''
    只留下re.sub裡面指定的東西:數字英文以及特殊指令
    '''
    # 只保留sub裡面寫的代號
    string = re.sub(r'[^A-Za-z0-9()./\-_:<>\\,"]', " ", string)
    string = string.replace('\\\\','\\')
    return string
# 要放入想測試的端口
def stringremove(item1):
    '''
    將特殊但是不需要的string內容刪掉避免影響指令輸入的正確性
    '''
    item2=f'{item1}'
    item2=item2.replace('[','').replace(']','')
    item2=item2.replace("'","")
    return item2
def local_json_to_firebase():
    '''
    res/firebase.json 上傳到 firebase
    '''
    try:
        with open(python_path+'/res/firebase.json') as f:
            firebase_local = json.load(f)
    except:
        print('沒有json檔')
    for item in firebase_local:
        try:
            db.child(item).update(firebase_local[item])
        except:
            print(item,'上傳失敗')
def firebase_to_local_json():
    '''
    firebase資料下載到本地的res/firebase.json
    '''
    #有網路的情況下
    try:
        full_data={}
        temp_data={}
        monitor=db.child('monitor').get()
        for item in monitor.each():
            temp_data[item.key()]=item.val()
        full_data['monitor']=temp_data
        temp_data={}
        user=db.child('user').get()
        for item in user.each():
            temp_data[item.key()]=item.val()
        full_data['user']=temp_data
        temp_data={}
        time=db.child('time').get()
        for item in time.each():
            temp_data[item.key()]=item.val()
        full_data['time']=temp_data
        json_object = json.dumps(full_data, indent = 4)
        # Writing to sample.json
        with open(python_path+"/res/firebase.json", "w") as outfile:
            outfile.write(json_object)
    except:
        print('沒有連上網路')
#輸入這次要跑的檔名
def getdict(file):
    '''
    基礎的指令碼搜集
    原理：以General為辨識基礎來區分第幾題,然後一航航搜尋是否為指令
    指令特徵：
        1. Window:
        2. MacOS:
        以上如果有出現就代表後面為相應指令,需要被記錄
    '''
    df=pd.read_csv(file)
    ddct={}
    i=0
    count=0
    cmd=[]
    example=[]
    last_i=None
    while i<len(df):
        q=stringremove(df[i:i+1].Title.values)
        cmd1=stringremove(df[i:i+1].Cmd.values).replace('\\xa0',' ')
        ex=stringremove(df[i:i+1].Example.values).replace('\\xa0',' ')
        if q=="General":
            if last_i ==None:
                last_i = i
            else:
                while last_i<i:
                    cmd2=stringremove(df[last_i:last_i+1].Cmd2.values).replace('\\xa0','')
                    if cmd2!="nan":
                        if cmd2 !='':
                            cmd.append(cmd2)
                    last_i+=1
                last_i = i
            if count>0:
                ddct[count]=[cmd,example]
            count+=1
            cmd=[]
            example=[]
            if cmd1!="nan":
                if cmd1 !='':
                    cmd.append(cmd1)
            if ex!="nan":
                if ex !='':
                    example.append(ex)
        else:
            if cmd1!="nan":
                if cmd1 !='':
                    cmd.append(cmd1)
            if ex!="nan":
                if ex !='':
                    example.append(ex)
        if i==len(df)-1:
            if last_i ==None:
                last_i = i
            else:
                while last_i<i:
                    cmd2=stringremove(df[last_i:last_i+1].Cmd2.values).replace('\\xa0',' ')
                    if cmd2!="nan":
                        if cmd2 !='':
                            cmd.append(cmd2)
                    last_i+=1
                last_i = i
            if count>0:
                ddct[count]=[cmd,example]
            count+=1
            cmd=[]
            example=[]
            if cmd1!="nan":
                if cmd1 !='':
                    cmd.append(cmd1)
            if ex!="nan":
                if ex !='':
                    example.append(ex)
        i+=1
    #將每題以順序跑出
    i=1
    while i<=len(ddct):
        tasktitle=''
        exptitle=''
        cmdtitle=[]
        j=0
        while j<len(ddct[i][0]):
            tasktitle=tasktitle+ddct[i][0][j]+'\n'
            if version == 'macos':
                if 'MacOS:' in ddct[i][0][j]:
                    cmdline=ddct[i][0][j].replace('\\\\\\','')
                    cmdline=cmdline.replace('\\xa0',' ')
                    cmdline=cmdline.replace('\xa0',' ')
                    cmdline=cmdline.replace('MacOS: ','')
                    cmdline=cmdline.replace('MacOS:','')
                    cmdline = cmdline.lstrip()
                    cmdline=stringremove(cmdline)
                    cmdtitle.append(cmdline)
            else:
                if 'Windows:' in ddct[i][0][j]:
                    # print(ddct[i][0][j])
                    cmdline=ddct[i][0][j].replace('\\\\\\','')
                    cmdline=cmdline.replace('\\xa0',' ')
                    cmdline=cmdline.replace('\xa0',' ')
                    cmdline=cmdline.replace('Windows: ','')
                    cmdline=cmdline.replace('Windows:','')
                    cmdline = cmdline.lstrip()
                    cmdline=stringremove(cmdline)
                    cmdtitle.append(cmdline)
            j+=1
        j=0
        while j<len(ddct[i][1]):
            exptitle=exptitle+ddct[i][1][j]+'\n'
            j+=1
        ddct[i]=[tasktitle,exptitle,cmdtitle]
        i+=1
    return ddct

def get_computer_user():
    '''
    抓取電腦ip *不是網路ip,是電腦特殊ip
    '''
    return socket.gethostname()
def ltos(s):  
    str1 = ""  
    for ele in s:  
        str1 += ele   
    return str1 

#internet

def internet_on():
    '''
    辨識網路是否正常
    '''
    url = "https://docs.google.com/spreadsheets/d/1YXFqt-ZuYGPlWIzjEdGGzy3-Rcj6kUV3zfhi7tfWET0/edit#gid=1326724140"
    timeout = 5
    try:
        request = requests.get(url, timeout=timeout)
        return True
    except (requests.ConnectionError, requests.Timeout) as exception:
        return False

# subinputname =["HDMI","HDMI1","HDMI2","HDMI3","DP","DP1","DP2","USB-C","Thunderbolt","VGA","DVI"]
# presetname1 = ['Standard ', 'ComfortView', 'Multiscreen Match', 'sRGB', 'AdobeRGB', 'REC709', 'REC2020', 'DCI-P3', 'DisplayP3', 'Warm', 'Cool', '5000k', '5700k', '6500k', '7500k', '9300k', '10000k', 'Custom', 'DCI P3 D65 G2.4 L100', 'BT.709 D65 BT1886 L100', 'BT.2020 D65 BT1886 L100', 'sRGB D65 sRGB L250', 'Adobe RGB D65 G2.2 L250']
# presetname3 =['Adobe RGB D50 G2.2 L250', 'Native', 'User 1', 'User 2', 'User 3', 'Custom 1', 'Custom 2', 'Custom 3', 'Cal 1', 'Cal 2', 'DICOM', 'Game', 'Movie', 'OSD : FPSSW : FPS Game', 'OSD : RTSSW : RTS Game', 'OSD : RPGDDM SW : RPG Game', 'Game 1', 'Game 2', 'Game 3', 'OSD : MOBA / RTSSW : RTS Game', 'OSD : SPORTSSW : SPORTS Game', 'sRGB D65 sRGB L120', 'Adobe RGB D65 G2.2 L160', 'Adobe RGB D50 G2.2 L160']
# preset1 =['Standard', 'ComfortView', 'MultiscreenMatch', 'sRGB', 'AdobeRGB', 'REC709', 'REC2020', 'DCI-P3', 'DisplayP3', 'Warm', 'Cool', '5000k', '5700k', '6500k', '7500k', '9300k', '10000k', 'Custom', 'DCI-P3','BT709', 'BT2020', 'sRGB', 'AdobeRGB1']
# preset3 = ['AdobeRGB2', 'Native', 'User1', 'User2', 'User3', 'Custom1', 'Custom2', 'Custom3', 'CAL1', 'CAL2', 'DICOM', 'Game', 'Movie', 'FPS Game', 'RTS Game', 'RPG Game', 'Game1', 'Game2', 'Game3', 'RTS Game', 'SPORTS Game', 'sRGB', 'AdobeRGB1', 'AdobeRGB2']
# window=["pbp-2h-fill","pbp-2h","pbp","split","pbp-2h-37","pbp-2h-73","pbp-2h-2674","pbp-2h-7426","pbp-2h-2575","pbp-2h-7525","pbp-2h-6733","pbp-2h-3367","pbp-2h-82","pbp-2h-28"]
runbox=['PIP','PBP(2window)','PBP(3or4window)','Rotation','Uniformity Compensation','Debug','Volume']

'''
firebase 資料
'''
firebaseConfig = {
  "apiKey": "AIzaSyAGsJobo2zotM0YcLygWzx0qytgia0cJZM",
  "authDomain": "ddmsetup.firebaseapp.com",
  "databaseURL": "https://ddmsetup-default-rtdb.asia-southeast1.firebasedatabase.app",
  "projectId": "ddmsetup",
  "storageBucket": "ddmsetup.appspot.com",
  "messagingSenderId": "890080399354",
  "appId": "1:890080399354:web:ce46fc23c75101cf681e43",
  "measurementId": "G-FX2CHHXH4B"
}
firebase=pyrebase.initialize_app(firebaseConfig) 
db=firebase.database()
# 授權登入用
auth=firebase.auth()
storage=firebase.storage()
def monitor_list():
    '''
    讀取所有已經存在的monitor資料
    '''
    data=['---select---']
    try:
        monitor=db.child('monitor').get()
        for item in monitor.each():
        # print(item.val())
            data.append(item.key())
    except:
        with open('res/firebase.json') as f:
            firebase_local = json.load(f)
        for item in firebase_local:
            if item=='monitor':
                for item2 in firebase_local[item]:
                    data.append(item2)
    return data

def get_sheet_list():
    '''
    讀取google 中的sheet
    '''
    try:
        auth_json_path =python_path+'/res/ddm-test-answer-2021-d7a5933c871b.json'
        gss_scopes = ['https://spreadsheets.google.com/feeds']
        #連線
        credentials = Credentials.from_service_account_file(auth_json_path,scopes=gss_scopes)
        gss_client = gspread.authorize(credentials)
        #開啟 Google Sheet 資料表
        spreadsheet_key = '1YXFqt-ZuYGPlWIzjEdGGzy3-Rcj6kUV3zfhi7tfWET0' 
        sheet = gss_client.open_by_key(spreadsheet_key)
        worksheet_list = sheet.worksheets()
        work=f'{worksheet_list}'
        work=list(work.split("'"))
        i=1
        worklist=['---select---']
        while i<len(work):
            worklist.append(work[i])
            i+=2
        return worklist
    except:
        print('沒網路,請連網！！')
class main(QMainWindow):
    '''
    整個GUI 內容, 用於設定ui功能以及ui位子
    '''
    def __init__(self):
        self.x=1500
        self.y=800
        super(main, self).__init__()
        self.setFixedSize(self.x,self.y)
        # 設置窗口標題
        self.setWindowTitle('LEO-commandline-auto')
        #應用的初始調色板
        self.origPalette = QApplication.palette()
        #外部輸入
        file = Path(python_path+'/res/firebase.json')
        if file.exists():
            os.remove(file)
        # local_json_to_firebase()
        firebase_to_local_json()
        self.subinputname_list=subinputname
        self.presetname1 = presetname1
        self.presetname3 = presetname3
        self.preset1_list=preset1
        self.preset3_list=preset3
        self.window_list=window
        self.skip_list=runbox
        times=['---select---','10','20','30','40','50']
        for i in range(10):
            if i>=1:
                times.append(str(60*i))
        self.time_list=times
        self.worklist=get_sheet_list()
        self.monitor=monitor_list()
        self.db=db
        # try:
        #     with open(python_path+'/res/firebase.json') as f:
        #         self.firebase_local = json.load(f)
        # except:
        #     print('無網路請聯網')
        # 初始值 設為None是為了偵測是否有輸入
        self.command_start_list = ['DDPM','DDM']
        self.command_start_choose = 'DDPM'
        self.ddct=getdict(beta_name)
        self.change_mode_state='auto'
        self.error_count_for_loop=0
        self.mode=None
        self.start_task_num=None
        self.end_task_num=None
        self.time_choose_num=None
        self.sheet_choose=None
        self.monitor_choose=None
        self.second_monitor_choose=None
        self.temp_input_list=[]
        self.temp_window_list=[]
        self.temp_preset1_list=[]
        self.temp_preset3_list=[]
        self.temp_skip_list=[]
        self.temp_input_list2=[]
        self.temp_window_list2=[]
        self.temp_preset1_list2=[]
        self.temp_preset3_list2=[]
        self.temp_skip_list2=[]
        self.all_state_lineedit=[]
        self.all_state_checkbox_list=[]
        self.monitor_main_port_item=None
        self.monitor_second_port_item=None
        self.monitor_third_port_item=None
        self.script_iwantto_stop=True
        self.script_temp_iwantto_stop=None
        # 優先讀取此電腦上次設定(如果沒網路就讀取res/firebase.json)
        user=socket.gethostname().replace('.','').replace('-','')
        try:
            self.firebase_monitor=db.child('user').child(user).get().val()['monitor']
        except:
            try:
                self.firebase_monitor=self.firebase_local['user'][user]['monitor']
            except:
                self.firebase_monitor=None
        try:
            self.firebase_mode=db.child('user').child(user).get().val()['mode']
        except:
            try:
                self.firebase_mode=self.firebase_local['user'][user]['mode']
            except:
                self.firebase_mode=None
        try:
            self.firebase_start_task=db.child('user').child(user).get().val()['start_task']
        except:
            try:
                self.firebase_start_task=self.firebase_local['user'][user]['start_task']
            except:
                self.firebase_start_task=None
        try:
            self.firebase_end_task=db.child('user').child(user).get().val()['end_task']
        except:
            try:
                self.firebase_end_task=self.firebase_local['user'][user]['end_task']
            except:
                self.firebase_end_task=None
        try:
            self.firebase_sheet=db.child('user').child(user).get().val()['sheet']
        except:
            try:
                self.firebase_sheet=self.firebase_local['user'][user]['sheet']
            except:
                self.firebase_sheet=None
        try:
            self.firebase_command_start = db.child('user').child(user).get().val()['command_start']
        except:
            try:
                self.firebase_command_start=self.firebase_local['user'][user]['command_start']
            except:
                self.firebase_command_start=None
        try:
            self.firebase_time=db.child('user').child(user).get().val()['time']
        except:
            try:
                self.firebase_time=self.firebase_local['user'][user]['time']
            except:
                self.firebase_time=None
        try:
            self.firebase_second_monitor=db.child('user').child(user).get().val()['second_monitor']
        except:
            try:
                self.firebase_second_monitor=self.firebase_local['user'][user]['second_monitor']
            except:
                self.firebase_second_monitor=None
        try:
            self.firebase_main_port=db.child('user').child(user).get().val()['main_port']
        except:
            try:
                self.firebase_main_port=self.firebase_local['user'][user]['main_port']
            except:
                self.firebase_main_port=None
        try:
            self.firebase_second_port=db.child('user').child(user).get().val()['second_port']
        except:
            try:
                self.firebase_second_port=self.firebase_local['user'][user]['second_port']
            except:
                self.firebase_second_port=None
        try:
            self.firebase_third_port=db.child('user').child(user).get().val()['third_port']
        except:
            try:
                self.firebase_third_port=self.firebase_local['user'][user]['third_port']
            except:
                self.firebase_third_port=None
        # 模式設定
        self.top_set=0
        self.set_mode_title = QLabel("模式選擇：",self)
        self.set_mode_title.setGeometry(self.x*0.03, self.y*self.top_set,self.x*0.15, 30)
        # 切換手動自動
        if self.change_mode_state=='auto':
            self.change_mode_button=QPushButton('切換為手動',self)
            self.change_mode_button.setGeometry(self.x*0.1,self.y*self.top_set,self.x*0.1, 20)
            self.change_mode_button.setStyleSheet("background-color: rgb(255,165,0)")
            self.change_mode_button.clicked.connect(self.change_mode_def)
            self.change_mode_button.show()
        else:
            self.change_mode_button=QPushButton('切換為自動',self)
            self.change_mode_button.setGeometry(self.x*0.1,self.y*self.top_set,self.x*0.1, 20)
            self.change_mode_button.setStyleSheet("background-color: rgb(152,245,255)")
            self.change_mode_button.clicked.connect(self.change_mode_def)
            self.change_mode_button.show()
            
        # 手動自動選取
        if self.change_mode_state=='auto':
            self.auto_radio_button = QRadioButton("略過手動",self)
            self.auto_radio_button.setGeometry(self.x*0.05, self.y*(self.top_set+0.05),self.x*0.1, 30)
            self.auto_radio_button.clicked.connect(self.auto_mode)
            self.hand_radio_button = QRadioButton("不略過手動",self)
            self.hand_radio_button.setGeometry(self.x*0.14, self.y*(self.top_set+0.05),self.x*0.1, 30)
            self.hand_radio_button.clicked.connect(self.hand_mode)
            if self.firebase_mode !=None:
                if self.firebase_mode=='auto_mode':
                    self.auto_radio_button.setChecked(True)
                    self.hand_radio_button.setChecked(False)
                else:
                    self.auto_radio_button.setChecked(False)
                    self.hand_radio_button.setChecked(True)
                self.mode=self.firebase_mode
            try:
                self.upload_button.hide()
            except:
                pass
        else:
            pass
        ###########後面全部加0.05
        # 題目輸入
        if self.change_mode_state=='auto':
            if self.firebase_start_task !=None:
                self.start_task_line = QLineEdit(str(self.firebase_start_task),self)
                self.start_task_num=self.firebase_start_task
            else:
                self.start_task_line = QLineEdit('',self)
            self.start_task_line.setPlaceholderText("開始題目：")
            self.start_task_line.setGeometry(self.x*0.05, self.y*(self.top_set+0.05*2), self.x*0.15, 30)
            self.start_task_line.textChanged.connect(lambda: self.start_task())
            if self.firebase_end_task !=None:
                self.end_task_line = QLineEdit(str(self.firebase_end_task),self)
                self.end_task_num=self.firebase_end_task
            else:
                self.end_task_line = QLineEdit('',self)
            self.end_task_line.setPlaceholderText("結束題目：")
            self.end_task_line.setGeometry(self.x*0.05, self.y*(self.top_set+0.05*3), self.x*0.15, 30)
            self.end_task_line.textChanged.connect(lambda: self.end_task())
        else:
            self.command_task_line=QSpinBox(self)
            numberoflast = int(list(self.ddct)[-1])
            self.command_task_line.setRange(1, numberoflast)
            self.command_task_line.setSingleStep(1)
            self.command_task_line.setWrapping(True)
            self.command_task_line.setValue(1)
            # self.command_task_line.setPlaceholderText("輸入題號：")
            self.command_task_line.setGeometry(self.x*0.03, self.y*(self.top_set+0.05), self.x*0.15, 30)
            self.command_task_line.valueChanged.connect(lambda: self.command_task())
            # upload button 
            self.upload_button=QPushButton('上傳',self)
            self.upload_button.setGeometry(self.x*0.18, self.y*(self.top_set+0.05), self.x*0.06, 30)
            self.upload_button.clicked.connect(lambda: self.upload())
            self.upload_button.show()
        #  ddm or ddpm
        self.label_command_start = QLabel("指令前綴選擇:",self)
        self.label_command_start.setGeometry(self.x*0.03, self.y*(self.top_set+0.05*4), self.x*0.15, 30)
        self.command_combo_box = QComboBox(self)
        self.command_combo_box.show()
        self.command_combo_box.addItems(self.command_start_list)
        if self.firebase_command_start!=None:
            try:
                self.command_combo_box.setCurrentIndex(self.command_start_list.index(str(self.firebase_command_start)))
                self.command_start_choose=self.firebase_command_start
            except:pass
        self.command_combo_box.currentIndexChanged.connect(self.command_start_choose_def)
        self.command_combo_box.setGeometry(self.x*0.05, self.y*(self.top_set+0.05*4.8), self.x*0.15, 30)
        # 時間選擇
        self.label_time = QLabel("選擇時間間隔：",self)
        self.label_time.setGeometry(self.x*0.03, self.y*(self.top_set+0.05*5.5), self.x*0.15, 30)
        self.time_combo_box = QComboBox(self)
        self.time_combo_box.show()
        self.time_combo_box.addItems(self.time_list)
        if self.firebase_time!=None:
            try:
                self.time_combo_box.setCurrentIndex(self.time_list.index(str(self.firebase_time)))
                self.time_choose_num=self.firebase_time
            except:pass
        self.time_combo_box.currentIndexChanged.connect(self.time_choose)
        self.time_combo_box.setGeometry(self.x*0.05, self.y*(self.top_set+0.05*6.3), self.x*0.15, 30)
        # 偵測可用的sheet
        self.label_sheet = QLabel("選擇sheet：",self)
        self.label_sheet.setGeometry(self.x*0.03, self.y*(self.top_set+0.05*7), self.x*0.15, 30)
        self.sheet = QComboBox(self)
        self.sheet.show()
        print(self.worklist)
        try:
            self.sheet.addItems(self.worklist)
        except:
            self.worklist=[self.firebase_sheet]
            self.sheet.addItems(self.worklist)
        self.sheet.currentIndexChanged.connect(self.save_sheet)
        self.sheet.setEditable(True)
        if self.firebase_sheet!=None:
            try:
                self.sheet.setCurrentIndex(self.worklist.index(self.firebase_sheet))
                self.sheet_choose=self.firebase_sheet
            except:pass
        self.sheet.setGeometry(self.x*0.05, self.y*(self.top_set+0.05*8), self.x*0.15, 30)
        self.create_sheet=QPushButton('建立新sheet', self)
        self.create_sheet.setGeometry(self.x*0.05,self.y*(self.top_set+0.05*9),self.x*0.15, 30)
        self.create_sheet.clicked.connect(self.create_new_sheet_to_qlineedit)
        # edit button
        self.edit_monitor_state=QPushButton('Edit',self)
        self.edit_monitor_state.setGeometry(self.x*0.27,self.y*(self.top_set+0.05*12),self.x*0.05, 30)
        self.edit_monitor_state.setStyleSheet("background-color: rgb(255,106,106)")
        self.edit_monitor_state.clicked.connect(self.edit_state)
        self.edit_monitor_state.hide()
        # 螢幕選擇
        self.label_monitor = QLabel('選擇monitor：',self)
        self.label_monitor.setGeometry(self.x*0.03, self.y*(self.top_set+0.05*10), self.x*0.2, 30)
        self.choose_monitor_label = QLabel('main：',self)
        self.choose_monitor_label.setGeometry(self.x*0.04, self.y*(self.top_set+0.05*10.9), self.x*0.15, 30)
        self.choose_monitor = QComboBox(self)
        self.choose_monitor.addItems(self.monitor)
        self.choose_monitor.currentIndexChanged.connect(self.save_monitor)
        self.choose_monitor.setEditable(True)
        if self.firebase_monitor!=None:
            try:
                self.choose_monitor.setCurrentIndex(self.monitor.index(self.firebase_monitor))
                self.monitor_choose=self.firebase_monitor
            except:pass
        self.choose_monitor.setGeometry(self.x*0.1, self.y*(self.top_set+0.05*11), self.x*0.15, 30)
        self.create_new_monitor_button=QPushButton('建立新monitor', self)
        self.create_new_monitor_button.setGeometry(self.x*0.1,self.y*(self.top_set+0.05*13),self.x*0.15, 30)
        self.create_new_monitor_button.clicked.connect(self.create_new_monitor_to_qlineedit)
        # 附螢幕
        self.choose_second_monitor_label = QLabel('second：',self)
        self.choose_second_monitor_label.setGeometry(self.x*0.04, self.y*(self.top_set+0.05*12), self.x*0.15, 30)
        self.choose_second_monitor = QComboBox(self)
        self.choose_second_monitor.addItems(self.monitor)
        self.choose_second_monitor.currentIndexChanged.connect(self.save_second_monitor)
        self.choose_second_monitor.setEditable(True)
        if self.firebase_second_monitor!=None:
            try:
                self.choose_second_monitor.setCurrentIndex(self.monitor.index(self.firebase_second_monitor))
                self.second_monitor_choose=self.firebase_second_monitor
            except:pass
        self.choose_second_monitor.setGeometry(self.x*0.1, self.y*(self.top_set+0.05*12), self.x*0.15, 30)
        # 主要輸入端選擇
        self.temp_input_port_list=self.subinputname_list
        self.temp_input_port_list.insert(0,'')
        self.label_monitor_main_port = QLabel('主螢幕輸入端口：',self)
        self.label_monitor_main_port.setGeometry(self.x*0.03, self.y*(self.top_set+0.05*14), self.x*0.2, 30)
        self.monitor_main_port = QComboBox(self)
        self.monitor_main_port.show()
        self.monitor_main_port.addItems(self.temp_input_port_list)
        self.monitor_main_port.currentIndexChanged.connect(self.save_monitor_main_port)
        if self.firebase_main_port!=None:
            try:
                self.monitor_main_port.setCurrentIndex(self.temp_input_port_list.index(self.firebase_main_port))
                self.monitor_main_port_item=self.firebase_main_port
            except:pass
        self.monitor_main_port.setGeometry(self.x*0.16, self.y*(self.top_set+0.05*14), self.x*0.15, 30)
        #
        self.label_monitor_second_port = QLabel('副螢幕輸入端口：',self)
        self.label_monitor_second_port.setGeometry(self.x*0.03, self.y*(self.top_set+0.05*15), self.x*0.2, 30)
        self.monitor_second_port = QComboBox(self)
        self.monitor_second_port.show()
        self.monitor_second_port.addItems(self.temp_input_port_list)
        self.monitor_second_port.currentIndexChanged.connect(self.save_monitor_second_port)
        if self.firebase_second_port!=None:
            try:
                self.monitor_second_port.setCurrentIndex(self.temp_input_port_list.index(self.firebase_second_port))
                self.monitor_second_port_item=self.firebase_second_port
            except:pass
        self.monitor_second_port.setGeometry(self.x*0.16, self.y*(self.top_set+0.05*15), self.x*0.15, 30)
        #
        self.label_monitor_third_port = QLabel('主螢幕輸入端口2：',self)
        self.label_monitor_third_port.setGeometry(self.x*0.03, self.y*(self.top_set+0.05*16), self.x*0.2, 30)
        self.monitor_third_port = QComboBox(self)
        self.monitor_third_port.show()
        self.monitor_third_port.addItems(self.temp_input_port_list)
        self.monitor_third_port.currentIndexChanged.connect(self.save_monitor_third_port)
        if self.firebase_third_port!=None:
            try:
                self.monitor_third_port.setCurrentIndex(self.temp_input_port_list.index(self.firebase_third_port))
                self.monitor_third_port_item=self.firebase_third_port
            except:pass
        self.monitor_third_port.setGeometry(self.x*0.16, self.y*(self.top_set+0.05*16), self.x*0.15, 30)

        # 輸入Montior_Services Tag
        self.Montior_Tag1_label = QLabel("選擇時間間隔：",self)
        self.Montior_Tag1_label.setGeometry(self.x*0.03, self.y*(self.top_set+0.85), self.x*0.15, 30)
        self.Montior_Tag1 = QLineEdit('',self)
        self.Montior_Tag1.setPlaceholderText("Monitor1_Services Tag")
        self.Montior_Tag1.setGeometry(self.x*0.1, self.y*(self.top_set+0.85), self.x*0.15, 30)
        self.Montior_Tag1.textChanged.connect(lambda: self.getMonitor_Tag1())

        # 輸入Montior_Services Tag
        self.Montior_Tag2_label = QLabel("選擇時間間隔：",self)
        self.Montior_Tag2_label.setGeometry(self.x*0.03, self.y*(self.top_set+0.9), self.x*0.15, 30)
        self.Montior_Tag2 = QLineEdit('',self)
        self.Montior_Tag2.setPlaceholderText("Monitor2_Services Tag")
        self.Montior_Tag2.setGeometry(self.x*0.1, self.y*(self.top_set+0.9), self.x*0.15, 30)
        self.Montior_Tag2.textChanged.connect(lambda: self.getMonitor_Tag2())


        # 開始按鈕
        self.startbutton=QPushButton('開始', self)
        self.startbutton.setIcon(self.style().standardIcon(QStyle.SP_MediaPlay))
        self.startbutton.setGeometry(self.x*0.13,self.y*0.95, self.x*0.1, 40)
        self.startbutton.clicked.connect(self.start)  
        # 暫停按鈕
        self.pause_button=QPushButton('暫停',self)
        self.pause_button.setIcon(self.style().standardIcon(QStyle.SP_MediaPause))
        self.pause_button.setGeometry(self.x*0.13,self.y*0.95, self.x*0.1, 40)
        self.pause_button.clicked.connect(self.pause)
        self.pause_button.hide()
        # 結束
        self.end_button=QPushButton('結束',self)
        self.end_button.setIcon(self.style().standardIcon(QStyle.SP_MediaStop))
        self.end_button.setGeometry(self.x*0.23,self.y*0.95, self.x*0.1, 40)
        self.end_button.clicked.connect(self.stop)
        self.end_button.hide()
        # keep going
        self.keep_going_button=QPushButton('繼續',self)
        self.keep_going_button.setIcon(self.style().standardIcon(QStyle.SP_MediaPlay))
        self.keep_going_button.setGeometry(self.x*0.13,self.y*0.95, self.x*0.1, 40)
        self.keep_going_button.clicked.connect(self.keep_going)
        self.keep_going_button.hide()
        # 偵測電腦
        self.detect_button=QPushButton('偵測型號',self)
        # self.detect_button.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))
        self.detect_button.setGeometry(self.x*0.03,self.y*0.95, self.x*0.1, 40)
        self.detect_button.clicked.connect(self.ddmbutton)
        # googlesheet 複製
        self.googlesheet=QPushButton('googlesheet',self)
        self.googlesheet.setGeometry(self.x*0.85,self.y*0.95, self.x*0.1, 40)
        self.googlesheet.clicked.connect(self.open_webbrowser)
    # 偵測長度變化
    # upload
   
    # change mode button
    def change_mode_def(self):
        self.change_mode_button.hide()
        if self.change_mode_state=='auto':
            self.change_mode_state='hand'
        else:
            self.change_mode_state='auto'
        if self.change_mode_state=='auto':
            self.change_mode_button=QPushButton('切換為手動',self)
            self.change_mode_button.setGeometry(self.x*0.1,self.y*self.top_set,self.x*0.1, 20)
            self.change_mode_button.setStyleSheet("background-color: rgb(255,165,0)")
            self.change_mode_button.clicked.connect(self.change_mode_def)
            self.change_mode_button.show()
            if self.script_iwantto_stop==True:
                self.startbutton.show()
            else:
                if self.script_temp_iwantto_stop==True:
                    self.keep_going_button.show()
                    self.end_button.show()
                elif self.script_temp_iwantto_stop==False:
                    self.pause_button.show()
                    self.end_button.show()      
            try:
                self.warn_no_answer.hide()
                QApplication.processEvents()
            except:pass      
        else:
            self.change_mode_button=QPushButton('切換為自動',self)
            self.change_mode_button.setGeometry(self.x*0.1,self.y*self.top_set,self.x*0.1, 20)
            self.change_mode_button.setStyleSheet("background-color: rgb(152,245,255)")
            self.change_mode_button.clicked.connect(self.change_mode_def)
            self.change_mode_button.show()
            # 隱藏自動化button
            try:
                self.keep_going_button.hide()
            except:
                pass
            try:
                self.end_button.hide()
            except:
                pass
            try:
                self.pause_button.hide()
            except:
                pass
            try:
                self.startbutton.hide()
            except:
                pass
        if self.change_mode_state=='auto':
            self.auto_radio_button = QRadioButton("略過手動",self)
            self.auto_radio_button.setGeometry(self.x*0.05, self.y*(self.top_set+0.05),self.x*0.1, 30)
            self.auto_radio_button.clicked.connect(self.auto_mode)
            self.hand_radio_button = QRadioButton("不略過手動",self)
            self.hand_radio_button.setGeometry(self.x*0.14, self.y*(self.top_set+0.05),self.x*0.1, 30)
            self.hand_radio_button.clicked.connect(self.hand_mode)
            if self.firebase_mode !=None:
                if self.firebase_mode=='auto_mode':
                    self.auto_radio_button.setChecked(True)
                    self.hand_radio_button.setChecked(False)
                else:
                    self.auto_radio_button.setChecked(False)
                    self.hand_radio_button.setChecked(True)
                self.mode=self.firebase_mode
            self.auto_radio_button.show()
            self.hand_radio_button.show()
        else:
            self.auto_radio_button.hide()
            self.hand_radio_button.hide()
        ###########後面全部加0.05
        # 題目輸入
        if self.change_mode_state=='auto':
            # 隱藏
            try:
                self.command_task_combobox.hide()
            except:
                pass
            self.command_task_line.hide()
            # 顯示
            if self.firebase_start_task !=None:
                self.start_task_line = QLineEdit(str(self.firebase_start_task),self)
                self.start_task_num=self.firebase_start_task
            else:
                self.start_task_line = QLineEdit('',self)
            self.start_task_line.setPlaceholderText("開始題目：")
            self.start_task_line.setGeometry(self.x*0.05, self.y*(self.top_set+0.05*2), self.x*0.15, 30)
            self.start_task_line.textChanged.connect(lambda: self.start_task())
            self.start_task_line.show()
            if self.firebase_end_task !=None:
                self.end_task_line = QLineEdit(str(self.firebase_end_task),self)
                self.end_task_num=self.firebase_end_task
            else:
                self.end_task_line = QLineEdit('',self)
            self.end_task_line.setPlaceholderText("結束題目：")
            self.end_task_line.setGeometry(self.x*0.05, self.y*(self.top_set+0.05*3), self.x*0.15, 30)
            self.end_task_line.textChanged.connect(lambda: self.end_task())
            self.end_task_line.show()
            try:
                self.upload_button.hide()
            except:
                pass
        else:
            
            # 隱藏
            self.start_task_line.hide()
            self.end_task_line.hide()
            # 顯示
            self.command_task_line=QSpinBox(self)
            numberoflast = int(list(self.ddct)[-1])
            self.command_task_line.setRange(1, numberoflast)
            self.command_task_line.setSingleStep(1)
            self.command_task_line.setWrapping(True)
            self.command_task_line.setValue(1)
            # self.command_task_line.setPlaceholderText("輸入題號：")
            self.command_task_line.setGeometry(self.x*0.03, self.y*(self.top_set+0.05), self.x*0.15, 30)
            self.command_task_line.valueChanged.connect(lambda: self.command_task())
            self.command_task_line.show()
            if version != 'macos':
                os.system(terminal_cmd)
                time.sleep(3)
                self.typing(cmd_cd)
    # googlesheet
    def open_webbrowser(self):
        webbrowser.open('https://docs.google.com/spreadsheets/d/1YXFqt-ZuYGPlWIzjEdGGzy3-Rcj6kUV3zfhi7tfWET0/edit#gid=2059002133')
    # 模式
    def auto_mode(self):
        self.mode='auto_mode'
        # 選取後清空結束的題號
        self.end_task_line.hide()
        self.end_task_line = QLineEdit('',self)
        self.end_task_num=None
        self.end_task_line.setPlaceholderText("結束題目：")
        self.end_task_line.setGeometry(self.x*0.05, self.y*(self.top_set+0.05*3), self.x*0.15, 30)
        self.end_task_line.textChanged.connect(lambda: self.end_task())
        self.end_task_line.show()
        QApplication.processEvents()
    def hand_mode(self):
        self.mode='hand_mode'
        try:
            self.end_task_num=int(self.start_task_line.text())+1
        except:
            self.end_task_num=None
        # 選取後自動填入題好
        self.end_task_line.hide()
        self.end_task_line = QLineEdit(str(self.end_task_num),self)
        self.end_task_line.setPlaceholderText("結束題目：")
        self.end_task_line.setGeometry(self.x*0.05, self.y*(self.top_set+0.05*3), self.x*0.15, 30)
        self.end_task_line.textChanged.connect(lambda: self.end_task())
        self.end_task_line.show()
    def command_task(self):
        # 警告
        if self.time_choose_num==None or self.time_choose_num=='---select---':
            try:
                self.time_num_None.hide()
            except:
                pass
            self.time_num_None=QLabel('尚未選擇時間',self)
            self.time_num_None.setStyleSheet('color: red')
            self.time_num_None.setGeometry(self.x*0.15, self.y*(self.top_set+0.05*4), self.x*0.15, 30)
            self.time_num_None.show()
        else:
            try:
                self.time_num_None.hide()
            except:
                pass
            try:
                self.command_task_combobox.hide()
            except:
                pass
            dictlist=['--select--']
            try:
                self.command_task_num=int(self.command_task_line.text())
                print(self.command_task_num)
                i=self.command_task_num
                reverse_ddct =self.reverse_dict(self.ddct)
                temp_dictlist=reverse_ddct[i][2]
                for temp in temp_dictlist:
                    dictlist.append(temp)

            except:
                dictlist=[]
            print('dictlist',dictlist)
            self.command_task_combobox = QComboBox(self)
            self.command_task_combobox.show()
            self.command_task_combobox.setEditable(True)
            self.command_task_combobox.addItems(dictlist)
            self.command_task_combobox.currentIndexChanged.connect(self.command_task_combobox_runcmd)
            self.command_task_combobox.setGeometry(self.x*0.03, self.y*(self.top_set+0.05*2), self.x*0.2, 60)
            self.command_task_combobox.show()
    
            
    def command_task_combobox_runcmd(self):
        self.save_command_temp=self.command_task_combobox.currentText()
        print('自動輸入=',self.save_command_temp)
        if '--select--' in self.save_command_temp:
            pass
        else:
            print(self.save_command_temp)
            self.command_typing(self.save_command_temp)
    def command_typing(self,cmd):
        cmd = cmd.replace('?','')
        if self.command_start_choose == 'DDPM':
            cmd = cmd.replace('DDPM','DDM')
        else:
            cmd = cmd.replace('ddpm','ddm')
        cmd = clean(cmd)
        try:
            self.upload_button.hide()
        except Exception as e:
            pass
        try:
            self.warn_no_answer.hide()
        except Exception as e:
            pass
        self.temp_script_iwantto_stop=self.script_iwantto_stop
        self.temp_script_temp_iwantto_stop=self.script_temp_iwantto_stop
        self.script_iwantto_stop=False
        self.script_temp_iwantto_stop=False
        if version == 'macos':
            os.system(terminal_cmd)
            time.sleep(3)
            self.typing(cmd_cd)
        ans=[]
        time.sleep(3)
        i=self.datalen()
        while os.path.isfile(ddpm_ddm(cmd_logfile,self.command_start_choose))==False :
            time.sleep(1)
        if self.command_start_choose == 'DDM':
            cmd = cmd.replace('DDPM','DDM')
            cmd = cmd.replace('ddpm','ddm')
        else:
            cmd = cmd.replace('DDM','DDPM')
            cmd = cmd.replace('ddm','ddpm') 
        keyboard.write(cmd) 
        keyboard.press_and_release('enter')
        # getdata
        i2=self.datalen()
        count=0
        while i2<=i and count<5 :
            i2=self.datalen()
            time.sleep(1)
            count+=1
        f = open(ddpm_ddm(cmd_logfile,self.command_start_choose), 'r')
        if i<i2:
            j=1
            for line in f.readlines():
                if j>i-1:
                    ans.append(line)
                j+=1
            # upload button 
            self.upload_button=QPushButton('上傳',self)
            self.upload_button.setGeometry(self.x*0.18, self.y*(self.top_set+0.05), self.x*0.06, 30)
            self.upload_button.clicked.connect(lambda: self.upload(ans))
            self.upload_button.show()    
            self.script_iwantto_stop=self.temp_script_iwantto_stop
            self.script_temp_iwantto_stop=self.temp_script_temp_iwantto_stop
            QApplication.processEvents()    
        else:
            ans=[]
            print('沒有輸入到terminal裡面')
            try:
                self.warn_no_answer.hide()
                QApplication.processEvents()
            except:pass
            self.warn_no_answer=QLabel('沒有找到答案',self)
            self.warn_no_answer.setStyleSheet('color: red')
            self.warn_no_answer.setGeometry(self.x*0.07, self.y*(self.top_set+0.05*3.3), self.x*0.2, 30)
            self.warn_no_answer.show()
        
    def upload(self,ans):
        if self.sheet_choose==None or self.sheet_choose=='---select---':
            try:
                self.sheet_choose_None.hide()
                QApplication.processEvents()
            except:pass
            self.sheet_choose_None=QLabel('尚未選擇sheet',self)
            self.sheet_choose_None.setStyleSheet('color: red')
            self.sheet_choose_None.setGeometry(self.x*0.15, self.y*(self.top_set+0.05*6), self.x*0.15, 30)
            self.sheet_choose_None.show()
        else:
            try:
                self.sheet_choose_None.hide()
                QApplication.processEvents()
            except:pass
            try:
                auth_json_path =python_path+'/res/ddm-test-answer-2021-d7a5933c871b.json'
                gss_scopes = ['https://spreadsheets.google.com/feeds']
                #連線
                #credentials = ServiceAccountCredentials.from_json_keyfile_name(auth_json_path,gss_scopes)
                credentials = Credentials.from_service_account_file(auth_json_path,scopes=gss_scopes)
                gss_client = gspread.authorize(credentials)
                #開啟 Google Sheet 資料表
                spreadsheet_key = '1YXFqt-ZuYGPlWIzjEdGGzy3-Rcj6kUV3zfhi7tfWET0' 
                googlename=self.sheet_choose
                print(self.sheet_choose)
                sheet = gss_client.open_by_key(spreadsheet_key).worksheet(googlename)
                df = pd.DataFrame(sheet.get_all_records())
                i=0
                check_success=0
                while i<len(df):
                    print(df.steporder[i],df.Comment[i])
                    if df.steporder[i]==self.command_task_num:
                        # values = [df.steporder[i],df.document[i],df.example[i],df.terminal[i],df.result[i],'',ans]
                        # sheet.insert_row(values,i) 
                        final_ans=df.Comment[i]+'\n'+ltos(ans)
                        sheet.update('G'+str(i+2),final_ans)
                        check_success+=1
                    i+=1
                print('check_success',check_success)
                if check_success==0:
                    #測試ddct[i][0]
                    try:
                        values = [self.command_task_num,self.ddct[self.command_task_num][0],self.ddct[self.command_task_num][1],'','','',ltos(ans)]
                        sheet.insert_row(values, len(sheet.get_all_values())+1) 
                    except Exception as e:
                        print(e)
            except Exception as e:
                print(e)
                print('沒連上網路,無法上傳到googlesheet')
            try:
                self.upload_button.hide()
            except Exception as e:
                print(e)
        
        # print('上傳',ans)
    def start_task(self):
        try:
            self.start_task_num=int(self.start_task_line.text())
        except:
            self.start_task_num=None
        if self.mode=='hand_mode':
            try:
                self.end_task_num=int(self.start_task_line.text())+1
            except:
                self.end_task_num=None
            self.end_task_line.hide()
            self.end_task_line = QLineEdit(str(self.end_task_num),self)
            self.end_task_line.setPlaceholderText("結束題目：")
            self.end_task_line.setGeometry(self.x*0.05, self.y*(self.top_set+0.05*3), self.x*0.15, 30)
            self.end_task_line.textChanged.connect(lambda: self.end_task())
            self.end_task_line.show()
        QApplication.processEvents()
        print('mode',self.mode,self.end_task_num)
    def getMonitor_Tag1(self):
        try:
            self.monitor_tag1_cmd=self.Montior_Tag1.text()
        except:
            self.monitor_tag1_cmd=None
    def getMonitor_Tag2(self):
        try:
            self.monitor_tag2_cmd=self.Montior_Tag2.text()
        except:
            self.monitor_tag2_cmd=None
    def end_task(self):
        try:
            self.end_task_num=int(self.end_task_line.text())
        except:
            self.end_task_num=None
    # 時間
    # self.time_combo_box = QComboBox(self)
    def time_choose(self):
        try:
            self.time_choose_num=int(self.time_combo_box.currentText())
        except:
            self.time_choose_num=None
    def command_start_choose_def(self):
        try:
            self.command_start_choose = self.command_combo_box.currentText()
        except:
            self.command_start_choose = 'DDM'
    # sheet 選擇
    # self.sheet = QComboBox(self)
    def save_sheet(self):
        self.sheet_choose=self.sheet.currentText()
    # self.create_sheet=QPushButton('建立新sheet', self)
    def create_new_sheet_to_qlineedit(self):
        # 隱藏原始Qcombobox
        self.label_sheet.hide()
        self.sheet.hide()
        self.create_sheet.hide()
        #建立填寫的表單
        self.label_new_sheet = QLabel("建立新sheet:",self)
        self.label_new_sheet.setGeometry(self.x*0.03, self.y*(self.top_set+0.05*7), self.x*0.15, 30)
        self.set_new_sheet = QLineEdit('',self)
        self.set_new_sheet.setPlaceholderText("輸入新sheet：")
        self.set_new_sheet.setGeometry(self.x*0.05, self.y*(self.top_set+0.05*8), self.x*0.15, 30)
        self.set_new_sheet.textChanged.connect(lambda: self.new_sheet_name())
        self.create_sheet=QPushButton('建立', self)
        self.create_sheet.setGeometry(self.x*0.05,self.y*(self.top_set+0.05*9),self.x*0.15, 30)
        self.create_sheet.clicked.connect(self.create_new_sheet)
        self.label_new_sheet.show()
        self.set_new_sheet.show()
        self.create_sheet.show()
        QApplication.processEvents()
    # self.set_new_sheet.setPlaceholderText("輸入新sheet：")
    def new_sheet_name(self):
        data=self.set_new_sheet.text()
        return data
    # self.create_sheet=QPushButton('建立', self)
    def create_new_sheet(self):
        self.label_new_sheet.hide()
        self.set_new_sheet.hide()
        self.create_sheet.hide()
        if self.new_sheet_name() not in self.worklist and self.new_sheet_name()!='':
            new_list=self.worklist
            new_list.insert(0,self.new_sheet_name())
        else:
            new_list=self.worklist
        self.sheet = QComboBox(self)
        self.sheet_choose=self.new_sheet_name()
        self.sheet.show()
        print(new_list)
        self.label_sheet = QLabel("選擇sheet:",self)
        self.label_sheet.setGeometry(self.x*0.03, self.y*(self.top_set+0.05*7), self.x*0.15, 30)
        self.sheet.addItems(new_list)
        self.sheet.currentIndexChanged.connect(self.save_sheet)
        self.sheet.setEditable(True)
        self.sheet.setGeometry(self.x*0.05, self.y*(self.top_set+0.05*8), self.x*0.15, 30)
        self.create_sheet=QPushButton('建立新sheet', self)
        self.create_sheet.setGeometry(self.x*0.05,self.y*(self.top_set+0.05*9),self.x*0.15, 30)
        self.create_sheet.clicked.connect(self.create_new_sheet_to_qlineedit)
        self.sheet.show()
        self.label_sheet.show()
        self.create_sheet.show()
        QApplication.processEvents()
    # 選擇螢幕
    # self.choose_monitor = QComboBox(self)
    def save_monitor(self):
        # 記錄現在的選項
        self.monitor_choose=self.choose_monitor.currentText()
        if self.monitor_choose=='---select---':
            try:
                self.edit_monitor_state.hide()
            except:pass
        else:
            # 提供編輯
            self.edit_monitor_state.show()
        try:
            self.temp_input_list=db.child('monitor').child(self.monitor_choose).get().val()['input']
        except:
            try:
                self.temp_input_list=self.firebase_local['monitor'][self.monitor_choose]['input']
            except:
                self.temp_input_list=[]
        print('input_list=',self.temp_input_list,self.monitor_choose)
        try:
            self.temp_window_list=db.child('monitor').child(self.monitor_choose).get().val()['window']
        except:
            try:
                self.temp_window_list=self.firebase_local['monitor'][self.monitor_choose]['window']
            except:
                self.temp_window_list=[]
        print('window_list=',self.temp_window_list,self.monitor_choose)
        try:
            self.temp_preset1_list=db.child('monitor').child(self.monitor_choose).get().val()['preset1']
        except:
            try:
                self.temp_preset1_list=self.firebase_local['monitor'][self.monitor_choose]['preset1']
            except:
                self.temp_preset1_list=[]
        print('preset1_list=',self.temp_preset1_list,self.monitor_choose)
        try:
            self.temp_preset3_list=db.child('monitor').child(self.monitor_choose).get().val()['preset3']
        except:
            try:
                self.temp_preset3_list=self.firebase_local['monitor'][self.monitor_choose]['preset3']
            except:
                self.temp_preset3_list=[]
        print('preset3_list=',self.temp_preset3_list,self.monitor_choose)
        try:
            self.temp_skip_list=db.child('monitor').child(self.monitor_choose).get().val()['skip']
        except:
            try:
                self.temp_skip_list=self.firebase_local['monitor'][self.monitor_choose]['skip']
            except:
                self.temp_skip_list=[]
        temp_new_input_list = []
        for i in self.temp_input_list:
            if i != None:
                temp_new_input_list.append(i)
        self.temp_input_list = temp_new_input_list
        temp_new_window_list = []
        for i in self.temp_window_list:
            if i != None:
                temp_new_window_list.append(i)
        self.temp_window_list = temp_new_window_list
        temp_new_preset1_list = []
        for i in self.temp_preset1_list:
            if i != None:
                temp_new_preset1_list.append(i)
        self.temp_preset1_list = temp_new_preset1_list
        temp_new_preset3_list = []
        for i in self.temp_preset3_list:
            if i != None:
                temp_new_preset3_list.append(i)
        self.temp_preset3_list = temp_new_preset3_list
        temp_new_skip_list = []
        for i in self.temp_skip_list:
            if i != None:
                temp_new_skip_list.append(i)
        self.temp_skip_list = temp_new_skip_list
        print('skip_list=',self.temp_skip_list,self.monitor_choose)
        # 標題
        try:
            self.label_input_checkbox.hide()
        except:
            pass
        try:
            self.label_window_checkbox.hide()
        except:pass
        try:
            self.label_preset1_checkbox.hide()
        except:pass
        try:
            self.label_skip_checkbox.hide()
        except:pass
        self.label_input_checkbox = QLabel('選擇input：',self)
        self.label_input_checkbox.setGeometry(self.x*0.25, self.y*(self.top_set), self.x*0.2, 30)
        self.label_input_checkbox.show()
        self.label_window_checkbox = QLabel('選擇切割模式：',self)
        self.label_window_checkbox.setGeometry(self.x*0.35, self.y*(self.top_set), self.x*0.2, 30)
        self.label_window_checkbox.show()
        self.label_preset1_checkbox = QLabel('選擇preset mode：',self)
        self.label_preset1_checkbox.setGeometry(self.x*0.48, self.y*(self.top_set), self.x*0.2, 30)
        self.label_preset1_checkbox.show()
        self.label_skip_checkbox = QLabel('選擇不跑的task：',self)
        self.label_skip_checkbox.setGeometry(self.x*0.8, self.y*(self.top_set), self.x*0.2, 30)
        self.label_skip_checkbox.show()
        # 顯示已有的項目
        for item in self.all_state_lineedit:
            item.hide()
        self.all_state_lineedit=[]
        count_check_box=1
        for i in self.temp_input_list:
            self.line_edit_input = QLabel(i,self)
            self.line_edit_input.setGeometry(self.x*0.25, self.y*(self.top_set+0.03*count_check_box), self.x*0.23, 30)
            self.all_state_lineedit.append(self.line_edit_input)
            self.line_edit_input.show()
            count_check_box+=1
        count_check_box=1
        for i in self.temp_window_list:
            self.line_edit_window = QLabel(i,self)
            self.line_edit_window.setGeometry(self.x*0.35, self.y*(self.top_set+0.03*count_check_box), self.x*0.23, 30)
            self.all_state_lineedit.append(self.line_edit_window)
            self.line_edit_window.show()
            count_check_box+=1
        count_check_box=1
        for i in self.temp_preset1_list:
            self.line_edit_preset1 = QLabel(i,self)
            self.line_edit_preset1.setGeometry(self.x*0.48, self.y*(self.top_set+0.03*count_check_box), self.x*0.23, 30)
            self.all_state_lineedit.append(self.line_edit_preset1)
            self.line_edit_preset1.show()
            count_check_box+=1
        count_check_box=1
        for i in self.temp_preset3_list:
            self.line_edit_preset3 = QLabel(i,self)
            self.line_edit_preset3.setGeometry(self.x*0.63, self.y*(self.top_set+0.03*count_check_box), self.x*0.23, 30)
            self.all_state_lineedit.append(self.line_edit_preset3)
            self.line_edit_preset3.show()
            count_check_box+=1
        count_check_box=1
        for i in self.temp_skip_list:
            self.line_edit_skip = QLabel(i,self)
            self.line_edit_skip.setGeometry(self.x*0.8, self.y*(self.top_set+0.03*count_check_box), self.x*0.23, 30)
            self.all_state_lineedit.append(self.line_edit_skip)
            self.line_edit_skip.show()
            count_check_box+=1
        # 顯示已有的設定
        QApplication.processEvents()
    # self.create_monitor=QPushButton('建立新monitor', self)
    # 要跟主螢幕變數區分
    def save_second_monitor(self):
        # 記錄現在的選項
        self.second_monitor_choose=self.choose_second_monitor.currentText()
        try:
            self.temp_input_list2=db.child('monitor').child(self.second_monitor_choose).get().val()['input']
        except:
            try:
                self.temp_input_list2=self.firebase_local['monitor'][self.second_monitor_choose]['input']
            except:
                self.temp_input_list2=[]
        try:
            self.temp_window_list2=db.child('monitor').child(self.second_monitor_choose).get().val()['window']
        except:
            try:
                self.temp_window_list2=self.firebase_local['monitor'][self.second_monitor_choose]['window']
            except:
                self.temp_window_list2=[]
        try:
            self.temp_preset1_list2=db.child('monitor').child(self.second_monitor_choose).get().val()['preset1']
        except:
            try:
                self.temp_preset1_list2=self.firebase_local['monitor'][self.second_monitor_choose]['preset1']
            except:
                self.temp_preset1_list2=[]
        try:
            self.temp_preset3_list2=db.child('monitor').child(self.second_monitor_choose).get().val()['preset3']
        except:
            try:
                self.temp_preset3_list2=self.firebase_local['monitor'][self.second_monitor_choose]['preset3']
            except:
                self.temp_preset3_list2=[]
        try:
            self.temp_skip_list2=db.child('monitor').child(self.second_monitor_choose).get().val()['skip']
        except:
            try:
                self.temp_skip_list2=self.firebase_local['monitor'][self.second_monitor_choose]['skip']
            except:
                self.temp_skip_list2=[]
    def create_new_monitor_to_qlineedit(self):
        # 隱藏原始Qcombobox
        self.label_monitor.hide()
        self.choose_monitor.hide()
        self.choose_second_monitor_label.hide()
        self.choose_second_monitor.hide()
        self.choose_monitor_label.hide()
        self.create_new_monitor_button.hide()
        for item in self.all_state_lineedit:
            item.hide()
        for item in self.all_state_checkbox_list:
            item.hide()
        try:
            self.save_monitor_edit.hide()
        except:pass
        try:
            self.edit_monitor_state.hide()
        except:pass
        # 建立填寫表單
        self.label_new_monitor = QLabel("建立新monitor:",self)
        self.label_new_monitor.setGeometry(self.x*0.03, self.y*(self.top_set+0.05*10), self.x*0.15, 30)
        self.set_new_monitor = QLineEdit('',self)
        self.set_new_monitor.setPlaceholderText("輸入新monitor：")
        self.set_new_monitor.setGeometry(self.x*0.05, self.y*(self.top_set+0.05*11), self.x*0.15, 30)
        self.set_new_monitor.textChanged.connect(lambda: self.new_monitor_name())
        self.create_monitor=QPushButton('建立', self)
        self.create_monitor.setGeometry(self.x*0.05,self.y*(self.top_set+0.05*12),self.x*0.15, 30)
        self.create_monitor.clicked.connect(self.create_new_monitor)
        self.label_new_monitor.show()
        self.set_new_monitor.show()
        self.create_monitor.show()
        QApplication.processEvents()
    # self.set_new_monitor.setPlaceholderText("輸入新monitor：")
    def save_monitor_main_port(self):
        self.monitor_main_port_item=self.monitor_main_port.currentText()
    def save_monitor_second_port(self):
        self.monitor_second_port_item=self.monitor_second_port.currentText()
    def save_monitor_third_port(self):
        self.monitor_third_port_item=self.monitor_third_port.currentText()
    def new_monitor_name(self):
        data=self.set_new_monitor.text()
        return data
    # self.create_monitor=QPushButton('建立', self)
    def create_new_monitor(self):
        self.label_new_monitor.hide()
        self.set_new_monitor.hide()
        self.create_monitor.hide()
        if self.new_monitor_name() not in self.monitor and self.new_monitor_name() !='':
            new_list=self.monitor
            new_list.insert(0,self.new_monitor_name())
        else:
            new_list=self.monitor
        #主螢幕
        self.monitor_choose=self.new_monitor_name()
        self.choose_monitor = QComboBox(self)
        print(new_list)
        self.label_monitor = QLabel("選擇monitor:",self)
        self.label_monitor.setGeometry(self.x*0.03, self.y*(self.top_set+0.05*10), self.x*0.2, 30)
        self.choose_monitor_label = QLabel('main：',self)
        self.choose_monitor_label.setGeometry(self.x*0.04, self.y*(self.top_set+0.05*10.9), self.x*0.1, 30)
        self.choose_monitor.addItems(new_list)
        self.choose_monitor.currentIndexChanged.connect(self.save_monitor)
        self.choose_monitor.setEditable(True)
        self.choose_monitor.setGeometry(self.x*0.1, self.y*(self.top_set+0.05*11), self.x*0.15, 30)
        self.create_new_monitor_button=QPushButton('建立新monitor', self)
        self.create_new_monitor_button.setGeometry(self.x*0.1,self.y*(self.top_set+0.05*13),self.x*0.15, 30)
        self.create_new_monitor_button.clicked.connect(self.create_new_monitor_to_qlineedit)
        # 附螢幕
        self.choose_second_monitor_label = QLabel('second：',self)
        self.choose_second_monitor_label.setGeometry(self.x*0.04, self.y*(self.top_set+0.05*12), self.x*0.2, 30)
        self.choose_second_monitor = QComboBox(self)
        self.choose_second_monitor.addItems(new_list)
        self.choose_second_monitor.currentIndexChanged.connect(self.save_second_monitor)
        self.choose_second_monitor.setEditable(True)
        self.choose_second_monitor.setGeometry(self.x*0.1, self.y*(self.top_set+0.05*12), self.x*0.15, 30)
        #
        if self.monitor_choose!='---select---' and self.monitor_choose!='':
            self.edit_monitor_state=QPushButton('Edit',self)
            self.edit_monitor_state.setGeometry(self.x*0.27,self.y*(self.top_set+0.05*12),self.x*0.05, 30)
            self.edit_monitor_state.setStyleSheet("background-color: rgb(255,106,106)")
            self.edit_monitor_state.clicked.connect(self.edit_state)
            self.edit_monitor_state.show()
            # 重設輸入端口
            self.temp_input_list=[]
            self.temp_window_list=[]
            self.temp_preset1_list=[]
            self.temp_preset3_list=[]
        else:
            try:
                self.edit_monitor_state.hide()
            except:pass
        self.create_new_monitor_button.show()
        self.label_monitor.show()
        self.choose_monitor_label.show()
        self.choose_monitor.show()
        self.save_monitor()
        self.choose_second_monitor_label.show()
        self.choose_second_monitor.show()
        QApplication.processEvents()
    # self.edit_monitor_state=QPushButton('Edit',self)
    def edit_state(self):
        self.choose_monitor.setEnabled(False)
        self.choose_second_monitor.setEnabled(False)
        self.create_new_monitor_button.hide()
        try:
            self.edit_monitor_state.hide()
        except:
            pass
        for item in self.all_state_lineedit:
            item.hide()
        # 記錄所有checkbox 地址
        self.all_state_checkbox_list=[]
        # input button
        count_check_box=1
        try:
            self.subinputname_list.remove('')
        except:pass
        print('temp_input_list=',self.temp_input_list)
        for i in self.subinputname_list:
            self.check_boxes_input = QCheckBox(i,self)
            if i in self.temp_input_list:
                self.check_boxes_input.setChecked(True)
            self.check_boxes_input.setGeometry(self.x*0.25, self.y*(self.top_set+0.03*count_check_box), self.x*0.23, 30)
            self.check_boxes_input.clicked.connect(self.click_input)
            self.all_state_checkbox_list.append(self.check_boxes_input)
            self.check_boxes_input.show()
            count_check_box+=1
        # window button
        count_check_box=1
        print('temp_window_list=',self.temp_window_list)
        for i in self.window_list:
            self.check_boxes_window = QCheckBox(i,self)
            if i in self.temp_window_list:
                self.check_boxes_window.setChecked(True)
            self.check_boxes_window.setGeometry(self.x*0.35, self.y*(self.top_set+0.03*count_check_box), self.x*0.23, 30)
            self.check_boxes_window.clicked.connect(self.click_window)
            self.all_state_checkbox_list.append(self.check_boxes_window)
            self.check_boxes_window.show()
            count_check_box+=1
        # 螢幕preset1
        count_check_box=1
        print('temp_preset1_list=',self.temp_preset1_list)
        for i in self.presetname1:
            index_num=self.presetname1.index(i)
            preset_data = self.preset1_list[index_num]
            self.check_boxes_preset1 = QCheckBox(i,self)
            if preset_data in self.temp_preset1_list:
                self.check_boxes_preset1.setChecked(True)
            self.check_boxes_preset1.setGeometry(self.x*0.48, self.y*(self.top_set+0.03*count_check_box), self.x*0.23, 30)
            self.check_boxes_preset1.clicked.connect(self.click_preset1)
            self.all_state_checkbox_list.append(self.check_boxes_preset1)
            self.check_boxes_preset1.show()
            count_check_box+=1
        count_check_box=1
        print('temp_preset3_list=',self.temp_preset3_list)
        for i in self.presetname3:
            index_num=self.presetname3.index(i)
            preset_data = self.preset3_list[index_num]
            self.check_boxes_preset3 = QCheckBox(i,self)
            if  preset_data in self.temp_preset3_list:
                self.check_boxes_preset3.setChecked(True)
            self.check_boxes_preset3.setGeometry(self.x*0.63, self.y*(self.top_set+0.03*count_check_box), self.x*0.23, 30)
            self.check_boxes_preset3.clicked.connect(self.click_preset3)
            self.all_state_checkbox_list.append(self.check_boxes_preset3)
            self.check_boxes_preset3.show()
            count_check_box+=1
        # 不跑的
        count_check_box=1
        print('temp_skip_list=',self.temp_skip_list)
        for i in self.skip_list:
            self.check_boxes_skip = QCheckBox(i,self)
            if i in self.temp_skip_list and i != None:
                self.check_boxes_skip.setChecked(True)
            self.check_boxes_skip.setGeometry(self.x*0.8, self.y*(self.top_set+0.03*count_check_box), self.x*0.23, 30)
            self.check_boxes_skip.clicked.connect(self.click_skip)
            self.all_state_checkbox_list.append(self.check_boxes_skip)
            self.check_boxes_skip.show()
            count_check_box+=1
        # 存取選項
        self.save_monitor_edit=QPushButton('save',self)
        self.save_monitor_edit.setGeometry(self.x*0.27,self.y*(self.top_set+0.05*12),self.x*0.05, 30)
        self.save_monitor_edit.setStyleSheet("background-color: rgb(255,215,0)")
        self.save_monitor_edit.clicked.connect(self.save_monitor_edit_answer)
        self.save_monitor_edit.show()
        QApplication.processEvents()
    # self.save_monitor_edit=QPushButton('save',self)
    def save_monitor_edit_answer(self):
        self.choose_monitor.setEnabled(True)
        self.choose_second_monitor.setEnabled(True)
        self.create_new_monitor_button.show()
        add_new_monitor_dict={}
        for item in self.all_state_checkbox_list:
            item.hide()
        self.save_monitor_edit.hide()
        self.edit_monitor_state.show()
        # 存入firebase
        try:
            self.db.child('monitor').child(self.monitor_choose).update({'input':self.temp_input_list})
            self.db.child('monitor').child(self.monitor_choose).update({'window':self.temp_window_list})
            self.db.child('monitor').child(self.monitor_choose).update({'preset1':self.temp_preset1_list})
            self.db.child('monitor').child(self.monitor_choose).update({'preset3':self.temp_preset3_list})
            self.db.child('monitor').child(self.monitor_choose).update({'skip':self.temp_skip_list})
        except:pass
        with open(python_path+'/res/firebase.json') as f:
            firebase_local = json.load(f)
        for item in firebase_local:
            if item=='monitor':
                firebase_local[item][self.monitor_choose]={'input':self.temp_input_list,
                                                            'window':self.temp_window_list,
                                                            'preset1':self.temp_preset1_list,
                                                            'preset3':self.temp_preset3_list,
                                                            'skip':self.temp_skip_list,
                                                            }
        json_object = json.dumps(firebase_local, indent = 4)
        # Writing to sample.json
        with open(python_path+"/res/firebase.json", "w") as outfile:
            outfile.write(json_object)
        with open(python_path+'/res/firebase.json') as f:
            self.firebase_local = json.load(f)
        self.save_monitor()
        QApplication.processEvents()
    # 顯示現在的狀態 （輸入端）,window,pbp,不跑的
    def click_input(self):
        print(self.temp_input_list)
        if self.sender().text() not in self.temp_input_list:
            self.temp_input_list.append(self.sender().text())
        else:
            self.temp_input_list.remove(self.sender().text())
        print(self.sender().text())
        print(self.temp_input_list)
    def click_window(self):
        print(self.temp_window_list)
        if self.sender().text() not in self.temp_window_list:
            self.temp_window_list.append(self.sender().text())
        else:
            self.temp_window_list.remove(self.sender().text())
        print(self.sender().text())
        print(self.temp_window_list)
    def click_preset1(self):
        print(self.temp_preset1_list)
        index_num=self.presetname1.index(self.sender().text())
        if self.preset1_list[index_num] not in self.temp_preset1_list:
            print('add into list preset1')
            self.temp_preset1_list.append(self.preset1_list[index_num])
        else:
            print('remove from list preset1')
            self.temp_preset1_list.remove(self.preset1_list[index_num])
        print(self.sender().text(),self.preset1_list[index_num])
        print(self.temp_preset1_list)
    def click_preset3(self):
        print(self.temp_preset3_list)
        index_num=self.presetname3.index(self.sender().text())
        if self.preset3_list[index_num] not in self.temp_preset3_list:
            print('add into list preset3')
            self.temp_preset3_list.append(self.preset3_list[index_num])
        else:
            print('remove from list preset3')
            self.temp_preset3_list.remove(self.preset3_list[index_num])
        print(self.sender().text(),self.preset3_list[index_num])
        print(self.temp_preset3_list)
    def click_skip(self):
        print(self.temp_skip_list)
        if self.sender().text() not in self.temp_skip_list:
            self.temp_skip_list.append(self.sender().text())
        else:
            self.temp_skip_list.remove(self.sender().text())
        print(self.sender().text())
        print(self.temp_skip_list)
    # 開關
    def turn_button(self,state):
        if state=='off':
            # 將關掉所有按鈕
            self.detect_button.hide()
            self.googlesheet.hide()
            self.start_task_line.setEnabled(False)
            self.end_task_line.setEnabled(False)
            self.auto_radio_button.setHidden(True)
            self.hand_radio_button.setHidden(True)
            try:
                self.edit_monitor_state.hide()
            except:
                pass
            self.create_new_monitor_button.hide()
            self.time_combo_box.setEnabled(False)
            self.sheet.setEnabled(False)
            try:
                self.set_new_sheet.setEnabled(False)
            except:
                pass
            self.create_sheet.hide()
            self.choose_monitor.setEnabled(False)
            self.choose_second_monitor.setEnabled(False)
            self.create_new_monitor_button.setEnabled(False)
            self.monitor_main_port.setEnabled(False)
            self.monitor_second_port.setEnabled(False)
            self.monitor_third_port.setEnabled(False)
            QApplication.processEvents()
            
        else:
            # 將開啟所有功能
            self.detect_button.show()
            self.googlesheet.show()
            self.start_task_line.setEnabled(True)
            self.end_task_line.setEnabled(True)
            self.auto_radio_button.setHidden(False)
            self.hand_radio_button.setHidden(False)
            try:
                self.edit_monitor_state.show()
            except:
                pass
            self.create_new_monitor_button.show()
            self.time_combo_box.setEnabled(True)
            self.sheet.setEnabled(True)
            try:
                self.set_new_sheet.setEnabled(True)
            except:pass
            self.create_sheet.show()
            self.choose_monitor.setEnabled(True)
            self.choose_second_monitor.setEnabled(True)
            self.create_new_monitor_button.setEnabled(True)
            self.monitor_main_port.setEnabled(True)
            self.monitor_second_port.setEnabled(True)
            self.monitor_third_port.setEnabled(True)
            QApplication.processEvents()

        
        
    # 開始結束按鈕
    # 還未加入螢幕內容,以及更新sheet 裡面的select選項過濾
    def start(self):
        # 偵測模式
        print('時間＝',self.time_choose_num)
        if self.mode==None:
            try:
                self.task_num_None.hide()
            except:
                pass
            self.task_num_None=QLabel('尚未選取模式',self)
            self.task_num_None.setStyleSheet('color: red')
            self.task_num_None.setGeometry(self.x*0.15, self.y*self.top_set, self.x*0.15, 30)
            self.task_num_None.show()
        else:
            # 偵測題號
            if self.start_task_num==None or self.end_task_num==None:
                try:
                    self.task_num_None.hide()
                except:pass
                if self.start_task_num==None and self.end_task_num==None:
                    self.task_num_None=QLabel('尚未填寫題號',self)
                elif self.start_task_num==None:
                    self.task_num_None=QLabel('尚未填寫開始題號',self)
                elif self.end_task_num==None:
                    self.task_num_None=QLabel('尚未填寫結束題號',self)
                self.task_num_None.setStyleSheet('color: red')
                self.task_num_None.setGeometry(self.x*0.11, self.y*self.top_set, self.x*0.15, 30)
                self.task_num_None.show()
            else:
                try:
                    self.task_num_None.hide()
                except:pass
        # 偵測時間
        if self.time_choose_num==None or self.time_choose_num=='---select---':
            try:
                self.time_num_None.hide()
            except:
                pass
            self.time_num_None=QLabel('尚未選擇時間',self)
            self.time_num_None.setStyleSheet('color: red')
            self.time_num_None.setGeometry(self.x*0.15, self.y*(self.top_set+0.05*4), self.x*0.15, 30)
            self.time_num_None.show()
        else:
            try:
                self.time_num_None.hide()
            except:
                pass
        # 偵測sheet
        if self.sheet_choose==None or self.sheet_choose=='---select---':
            try:
                self.sheet_choose_None.hide()
            except:pass
            self.sheet_choose_None=QLabel('尚未選擇sheet',self)
            self.sheet_choose_None.setStyleSheet('color: red')
            self.sheet_choose_None.setGeometry(self.x*0.15, self.y*(self.top_set+0.05*6), self.x*0.15, 30)
            self.sheet_choose_None.show()
        else:
            try:
                self.sheet_choose_None.hide()
            except:pass
        # 偵測螢幕輸入
        if self.monitor_choose==None or self.monitor_choose=='---select---':
            try:
                self.monitor_choose_None.hide()
            except:pass
            self.monitor_choose_None=QLabel('尚未選擇monitor',self)
            self.monitor_choose_None.setStyleSheet('color: red')
            self.monitor_choose_None.setGeometry(self.x*0.15, self.y*(self.top_set+0.05*9), self.x*0.15, 30)
            self.monitor_choose_None.show()
        else:
            try:
                self.monitor_choose_None.hide()
            except:pass
        if self.temp_input_list==[] or self.temp_window_list==[] or self.temp_preset1_list==[] or self.temp_preset3_list==[]:
            try:
                self.monitor_choose_None.hide()
            except:pass
            if self.temp_input_list==[]:
                self.monitor_choose_None=QLabel('尚未選取input',self)
            elif self.temp_window_list==[]:
                self.monitor_choose_None=QLabel('尚未選取window',self)
            elif self.temp_preset1_list==[] and self.temp_preset3_list==[]:
                self.monitor_choose_None=QLabel('尚未選取preset',self)
            try:                
                self.monitor_choose_None.setStyleSheet('color: red')
                self.monitor_choose_None.setGeometry(self.x*0.15, self.y*(self.top_set+0.05*9), self.x*0.15, 30)
                self.monitor_choose_None.show()
            except:
                pass
        if self.monitor_main_port_item==None or self.monitor_main_port_item=='' :
            try:
                self.monitor_main_port_item_None.hide()
            except:pass
            self.monitor_main_port_item_None=QLabel('尚未選擇主螢幕port',self)
            self.monitor_main_port_item_None.setStyleSheet('color: red')
            self.monitor_main_port_item_None.setGeometry(self.x*0.3, self.y*(self.top_set+0.05*13), self.x*0.15, 30)
            self.monitor_main_port_item_None.show()
        else:
            try:
                self.monitor_main_port_item_None.hide()
            except:
                pass
            
        QApplication.processEvents()
        # 都有輸入
        print(self.temp_input_list,self.temp_window_list,self.temp_preset1_list)
        if self.sheet_choose!=None and self.time_choose_num!=None and self.start_task_num!=None and self.end_task_num!=None and self.temp_input_list!=[] and self.temp_window_list!=[] and self.temp_preset1_list!=[] and self.mode!=None and self.monitor_main_port_item!=None and self.monitor_main_port_item!='':
            self.startbutton.hide()
            self.pause_button.show()
            self.end_button.show()
            # input button
            self.label_input_checkbox.hide()
            # window button
            self.label_window_checkbox.hide()
            # 螢幕preset
            self.label_preset1_checkbox.hide()
            # 錯過的題目
            self.label_skip_checkbox.hide()
            # 啟動
            self.script_iwantto_stop=False
            self.script_temp_iwantto_stop=False
            for item in self.all_state_checkbox_list:
                try:
                    item.hide()
                except:
                    pass
            for item in self.all_state_lineedit:
                try:
                    item.hide()
                except:
                    pass
            # 建立使用者設定 並且下次打開後回傳
            user=socket.gethostname().replace('.','').replace('-','')
            try:
                self.db.child('user').child(user).update({'sheet':self.sheet_choose,
                                                        'time':self.time_choose_num,
                                                        'start_task':self.start_task_num,
                                                        'end_task':self.end_task_num,
                                                        'monitor':self.monitor_choose,
                                                        'second_monitor':self.second_monitor_choose,
                                                        'main_port':self.monitor_main_port_item,
                                                        'second_port':self.monitor_second_port_item,
                                                        'third_port':self.monitor_third_port_item,
                                                        'mode':self.mode,
                                                        'command_start':self.command_start_choose,
                                                        })
            except:
                pass
            # 紀錄時間
            temp_time_record=str(dt2.now())
            temp_time_record=temp_time_record.split('.')[0]
            print(temp_time_record)
            try:
                self.db.child('time').child(self.monitor_choose).child(temp_time_record).update({'user':user,
                                                                                                'sheet':self.sheet_choose,
                                                                                                'second_monitor':self.second_monitor_choose,
                                                                                                'main_port':self.monitor_main_port_item,
                                                                                                'second_port':self.monitor_second_port_item,
                                                                                                'third_port':self.monitor_third_port_item,
                                                                                                })
            except:
                pass
            with open(python_path+'/res/firebase.json') as f:
                firebase_local = json.load(f)
            for item in firebase_local:
                if item=='user':
                    firebase_local[item][user]={'sheet':self.sheet_choose,
                                                'time':self.time_choose_num,
                                                'start_task':self.start_task_num,
                                                'end_task':self.end_task_num,
                                                'monitor':self.monitor_choose,
                                                'second_monitor':self.second_monitor_choose,
                                                'main_port':self.monitor_main_port_item,
                                                'second_port':self.monitor_second_port_item,
                                                'third_port':self.monitor_third_port_item,
                                                'mode':self.mode,
                                                'command_start':self.command_start_choose,
                                                }
                elif item=='time':
                    try:
                        firebase_local[item][self.monitor_choose][temp_time_record]={'user':user,
                                                                                    'sheet':self.sheet_choose,
                                                                                    'second_monitor':self.second_monitor_choose,
                                                                                    'main_port':self.monitor_main_port_item,
                                                                                    'second_port':self.monitor_second_port_item,
                                                                                    'third_port':self.monitor_third_port_item,
                                                                                    }
                    except:
                        firebase_local[item][self.monitor_choose]={temp_time_record:{'user':user,
                                                                                    'sheet':self.sheet_choose,
                                                                                    'second_monitor':self.second_monitor_choose,
                                                                                    'main_port':self.monitor_main_port_item,
                                                                                    'second_port':self.monitor_second_port_item,
                                                                                    'third_port':self.monitor_third_port_item,
                                                                                    }}
                    
            json_object = json.dumps(firebase_local, indent = 4)
            # Writing to sample.json
            with open(python_path+"/res/firebase.json", "w") as outfile:
                outfile.write(json_object)
            with open(python_path+'/res/firebase.json') as f:
                self.firebase_local = json.load(f)
            self.turn_button('off')
            try:
                self.change_mode_button.hide()
            except:
                pass
            QApplication.processEvents()
            self.my_script()
    def pause(self):
        self.script_temp_iwantto_stop=True
        self.keep_going_button.show()
        self.pause_button.hide()
        self.turn_button('on')
        self.change_mode_button.show()
        QApplication.processEvents()
    def keep_going(self):
        self.script_temp_iwantto_stop=False
        self.keep_going_button.hide()
        self.pause_button.show()
        self.turn_button('off')
        try:
            self.change_mode_button.hide()
        except:
            pass
        QApplication.processEvents() 
        time.sleep(5)
    def stop(self):
        self.change_mode_button.show()
        self.startbutton.show()
        try:
            self.pause_button.hide()
        except:pass
        try:
            self.keep_going_button.hide()
        except:pass
        self.end_button.hide()
        self.script_iwantto_stop=True
        self.label_input_checkbox.show()
        # window button
        self.label_window_checkbox.show()
        # 螢幕preset
        self.label_preset1_checkbox.show()
        # 錯過的題目
        self.label_skip_checkbox.show()
        self.save_monitor()
        self.turn_button('on')
        QApplication.processEvents()
    
    # 辨識模式
    def my_script(self):
        if self.mode=='auto_mode':
            self.start_mode('auto')
        else:
            self.start_mode('hand')
    def ddmbutton(self):
        os.system(terminal_cmd)
        time.sleep(3)
        self.script_iwantto_stop=False
        self.script_temp_iwantto_stop=False
        self.typing(cmd_cd)
        self.typing(cmd_getAttritube)
        keyboard.press_and_release('enter')
        self.script_iwantto_stop=True
        self.script_temp_iwantto_stop=True
    
    def no_reponse_typing(self,cmd):
        if self.command_start_choose == 'DDM':
            cmd = cmd.replace('DDPM','DDM')
            cmd = cmd.replace('ddpm','ddm')
        else:
            cmd = cmd.replace('DDM','DDPM')
            cmd = cmd.replace('ddm','ddpm') 
        keyboard.write(cmd) 
        keyboard.press_and_release('enter')
    # 自動化程式
    def typing(self,cmd):
        if self.command_start_choose == 'DDM':
            cmd = cmd.replace('DDPM','DDM')
            cmd = cmd.replace('ddpm','ddm')
        else:
            cmd = cmd.replace('DDM','DDPM')
            cmd = cmd.replace('ddm','ddpm') 
        t_end = time.time() + 1
        while time.time() < t_end:
            if self.script_iwantto_stop==True:
                return False
            elif self.script_temp_iwantto_stop==True:
                while self.script_temp_iwantto_stop==True:
                    QApplication.processEvents()
            QApplication.processEvents()
        keyboard.write(cmd) 
        keyboard.press_and_release('enter')
        t_end = time.time() + 5
        while time.time() < t_end:
            if self.script_iwantto_stop==True:
                return False
            elif self.script_temp_iwantto_stop==True:
                while self.script_temp_iwantto_stop==True:
                    QApplication.processEvents()
            QApplication.processEvents()
        return True
    def datalen(self):
        i=0
        file = Path(ddpm_ddm(cmd_logfile,self.command_start_choose))
        if file.exists():
            f = open(ddpm_ddm(cmd_logfile,self.command_start_choose), 'r')
            i=1
            for line in f.readlines():
                f=line
                i+=1
            pass
        else:
            print('產生ddm.txt檔案')
            temp1 = self.typing(cmd_cd)
            temp2 =self.typing(cmd_mkdir)
            temp3 = self.typing(cmd_logcreate)
            if temp1 == False or temp2 == False or temp3 == False:
                return False
            temp_out=0
            while os.path.isfile(ddpm_ddm(cmd_logfile,self.command_start_choose))==False  and temp_out<=self.time_choose_num:
                time.sleep(1)
                temp_out+=1
                if self.script_iwantto_stop==True:
                    return False
                elif self.script_temp_iwantto_stop==True:
                    while self.script_temp_iwantto_stop==True:
                        QApplication.processEvents()
                QApplication.processEvents()
            filetxt=Path(ddpm_ddm(cmd_logfile,self.command_start_choose))
            if filetxt.exists():
                f = open(ddpm_ddm(cmd_logfile,self.command_start_choose), 'r')
                i=1
                for line in f.readlines():
                    f=line
                    i+=1
            else:
                pass
        return i
    def getdata(self,i):
        ans=[]
        while os.path.isfile(ddpm_ddm(cmd_logfile,self.command_start_choose))==False :
            time.sleep(1)
            if self.script_iwantto_stop==True:
                return False
            elif self.script_temp_iwantto_stop==True:
                while self.script_temp_iwantto_stop==True:
                    QApplication.processEvents()
            QApplication.processEvents()
        i2=self.datalen()
        if i2 == False:
            return False
        self.error_count_for_loop = 0
        if '2:' in self.cmd and '1:' in self.cmd:
            temp_addi = 1
        else:
            temp_addi = 0
        
        while i+temp_addi >=i2 and self.error_count_for_loop<4:
            temp_out=0
            if self.error_count_for_loop>0 and i+temp_addi>=i2:
                '''
                預防放大鏡
                '''
                time.sleep(10)
                if version == 'macos':
                    os.system(terminal_cmd)
                else:
                    pyautogui.leftClick()
                self.typing(self.cmd)   
            
            while i+temp_addi>=i2 and temp_out<=self.time_choose_num:
                print('tempout=',temp_out,'timelimit=',self.time_choose_num)
                time.sleep(1)
                temp_out+=1
                i2=self.datalen()
                if i2 == False:
                    return False
                if self.script_iwantto_stop==True:
                    return False
                elif self.script_temp_iwantto_stop==True:
                    while self.script_temp_iwantto_stop==True:
                        QApplication.processEvents()
                QApplication.processEvents()
            if i<i2:
                break
            self.error_count_for_loop +=1
        f = open(ddpm_ddm(cmd_logfile,self.command_start_choose), 'r')
        if i<i2:
            j=1
            print('come from DDM.txt')
            for line in f.readlines():
                if j>i-1:
                    ans.append(line)
                j+=1
            self.error_count_for_loop=0
        else: 
            print('runcmd重複跑回傳=',self.cmd)
            ans.append("No response!!")
            self.error_count_for_loop=0 
        return ans

        
    #打完指令後等待答案出來的時間
    def runcmd(self,cmdans,cmd):
        cmd = cmd.replace('DDPM','DDM')
        cmd = cmd.replace('ddpm','ddm')
        cmd = cmd.replace('<Monitor_Service Tag>',str(self.monitor_tag1_cmd)).replace('<Monitor1_Service Tag>',str(self.monitor_tag1_cmd)).replace('<Monitor2_Service Tag>',str(self.monitor_tag2_cmd))
        cmd = clean(cmd)
        run_bool = False
        if version == 'macos':
            if './DDM' in cmd:
                run_bool = True
        else:
            if 'ddm' in cmd[0:5] or 'DDM' in cmd[0:5]:
                run_bool = True
        if run_bool == True:
            print('runcmd=',cmd)
            self.cmd=cmd
            QApplication.processEvents()
            errorrun=0
            errorcount=0
            firsti=self.datalen()
            temp = self.typing(cmd)
            q=self.getdata(firsti)
            if q == False or firsti == False or temp == False:
                return False
            print('cmd=',cmd,'cmdans=',q)
            i=0
            try:
                while i<len(q):
                    if 'Error' in q[i] or 'error' in q[i]:
                        errorrun=1
                    i+=1
                if errorrun==1:
                    i=0
                    while i<4:
                        firsti=self.datalen()
                        '''
                        預防放大鏡
                        '''
                        time.sleep(10)
                        self.typing(cmd)
                        q=self.getdata(firsti)
                        if q == False or firsti == False or temp == False:
                            return False
                        it=0
                        while it<len(q):
                            if 'Error' in q[it] or 'error' in q[it]:
                                errorcount=1
                            if it==len(q)-1:
                                if errorcount==0:
                                    i=5
                                    break
                            it+=1
                        i+=1
                i=0
                while i<len(q):
                    if q[i]=="No response!!":
                        try:
                            cmdans=cmdans+q[i]+"\n<br>\n"
                        except:
                            cmdans=q[i]+"\n<br>\n"
                    else:
                        try:
                            cmdans=cmdans+q[i]+"<br>\n"
                        except:
                            cmdans=q[i]+"<br>\n"
                   
                    i+=1
                '''
                預防放大鏡
                '''
                time.sleep(10)
                return cmdans
            except Exception as e:
                '''
                預防放大鏡
                '''
                time.sleep(10)
                print(e)
        else:
            print('no response command')
            '''
            預防放大鏡
            '''
            time.sleep(10)
            self.no_reponse_typing(cmd)
            return cmdans
    def reverse_dict(self,file2):
        # preset=self.temp_preset1_list+self.temp_preset3_list
        if self.monitor_main_port_item==None or self.monitor_main_port_item=='':
            subinput2="USB-C"
        else:
            subinput2=self.monitor_main_port_item
        preset = self.temp_preset1_list+self.temp_preset3_list
        window=self.temp_window_list
        subinputname=self.temp_input_list
        print('self.temp_input_list=',self.temp_input_list)
        for title in file2:
            new_commandlist = []
            for command in file2[title][2]:
                if '<' in command or '>' in command:
                    command = command.replace('<Monitor_Service Tag>',str(self.monitor_tag1_cmd)).replace('<Monitor1_Service Tag>',str(self.monitor_tag1_cmd)).replace('<Monitor2_Service Tag>',str(self.monitor_tag2_cmd))
                    if 'input' in command.lower() and '2:' not in command and command.count('<')>1:
                        cmdlist = command.split('<')
                        cmdfinal = []
                        temp_input = self.temp_input_list
                        count_sub = command.count('<')
                        input_count = command.lower().count('input')
                        if count_sub>=input_count:
                            count_sub = input_count
                        else:
                            pass
                        sub_start = 0
                        for cmd in cmdlist:
                            tempcmdfinal = cmdfinal
                            if '>' in cmd:
                                cmd = cmd.split('>')
                                if 'input' in cmd[0] or 'Input' in cmd[0]:
                                    cmdfinal = []
                                    if 'main' in cmd[0]:
                                        if tempcmdfinal != []:
                                            for tempcmd in tempcmdfinal:
                                                if tempcmd[-1] != ' ':
                                                    cmdfinal.append(tempcmd+' '+subinput2)
                                                else:
                                                    cmdfinal.append(tempcmd+subinput2)
                                        else:
                                            cmdfinal.append(subinput2)
                                        try:
                                            temp_input.remove(subinput2)
                                        except:pass
                                        count_sub-=1
                                    else:
                                        if tempcmdfinal != []:
                                            if count_sub>1 and len(tempcmdfinal)>1:
                                                count_n = 0
                                                for data in temp_input[sub_start:]:
                                                    temp_num = temp_input[sub_start:].index(data)
                                                    if temp_num%count_sub==0:
                                                        tempcmd = tempcmdfinal[count_n]
                                                        if tempcmd[-1] != ' ':
                                                            cmdfinal.append(tempcmd+' '+data)
                                                        else:
                                                            cmdfinal.append(tempcmd+data)
                                                        count_n+=1
                                            else:
                                                for data in temp_input[sub_start:]:
                                                    temp_num = temp_input[sub_start:].index(data)
                                                    if temp_num%count_sub==0:
                                                        for tempcmd in tempcmdfinal:
                                                            if tempcmd[-1] != ' ':
                                                                cmdfinal.append(tempcmd+' '+data)
                                                            else:
                                                                cmdfinal.append(tempcmd+data)
                                            sub_start+=1
                                if 'window' in cmd[0]:
                                    cmdfinal = []
                                    if tempcmdfinal != []:
                                        for data in window:
                                            for tempcmd in tempcmdfinal:
                                                if tempcmd[-1] != ' ':
                                                    cmdfinal.append(tempcmd+' '+data)
                                                else:
                                                    cmdfinal.append(tempcmd+data)
                                    else:
                                        for data in window:
                                            cmdfinal.append(data)
                                if 'preset' in cmd[0]:
                                    cmdfinal = []
                                    if tempcmdfinal != []:
                                        for data in preset:
                                            for tempcmd in tempcmdfinal:
                                                if tempcmd[-1] != ' ':
                                                    cmdfinal.append(tempcmd+' '+data)
                                                else:
                                                    cmdfinal.append(tempcmd+data)
                                    else:
                                        for data in preset:
                                            cmdfinal.append(data)
                                if len(cmd)>1:
                                    tempcmdfinal = cmdfinal
                                    cmdfinal = []
                                    for tempcmd in tempcmdfinal:
                                        cmdfinal.append(tempcmd+cmd[1])
                            else:
                                cmdfinal = []
                                if tempcmdfinal != []:
                                    for tempcmd in tempcmdfinal:
                                        if tempcmd[-1] != ' ':
                                            cmdfinal.append(tempcmd+' '+cmd)
                                        else:
                                            cmdfinal.append(tempcmd+cmd)
                                else:
                                    cmdfinal = [cmd]
                        for item in cmdfinal:
                            # print(item)
                            new_commandlist.append(item)

                    else:
                        '''
                        add
                        '''
                        cmdlist = command.split('<')
                        print('title=',title,'down',cmdlist)
                        cmdfinal = []
                        # print(cmdlist)
                        for cmd in cmdlist:
                            tempcmdfinal = cmdfinal
                            if '>' in cmd:
                                cmd = cmd.split('>')
                                if 'input' in cmd[0] or 'Input' in cmd[0]:
                                    cmdfinal = []
                                    if 'main' in cmd[0]:
                                        if tempcmdfinal != []:
                                            for tempcmd in tempcmdfinal:
                                                if tempcmd[-1] != ' ':
                                                    cmdfinal.append(tempcmd+''+subinput2)
                                                else:
                                                    cmdfinal.append(tempcmd+subinput2)
                                        else:
                                            cmdfinal.append(subinput2)
                                    else:
                                        if tempcmdfinal != []:
                                            for data in self.temp_input_list:
                                                for tempcmd in tempcmdfinal:
                                                    if tempcmd[-1] != ' ':
                                                        cmdfinal.append(tempcmd+''+data)
                                                    else:
                                                        cmdfinal.append(tempcmd+data)
                                        else:
                                            for data in self.temp_input_list:
                                                cmdfinal.append(data)
                                if 'window' in cmd[0]:
                                    cmdfinal = []
                                    if tempcmdfinal != []:
                                        for data in window:
                                            for tempcmd in tempcmdfinal:
                                                if tempcmd[-1] != ' ':
                                                    cmdfinal.append(tempcmd+' '+data)
                                                else:
                                                    cmdfinal.append(tempcmd+data)
                                    else:
                                        for data in window:
                                            cmdfinal.append(data)
                                if 'preset' in cmd[0]:
                                    cmdfinal = []
                                    if tempcmdfinal != []:
                                        for data in preset:
                                            for tempcmd in tempcmdfinal:
                                                if tempcmd[-1] != ' ':
                                                    cmdfinal.append(tempcmd+' '+data)
                                                else:
                                                    cmdfinal.append(tempcmd+data)
                                    else:
                                        for data in preset:
                                            cmdfinal.append(data)
                                if len(cmd)>1:
                                    tempcmdfinal = cmdfinal
                                    cmdfinal = []
                                    for tempcmd in tempcmdfinal:
                                        cmdfinal.append(tempcmd+cmd[1])
                            else:
                                cmdfinal = []
                                if tempcmdfinal != []:
                                    for tempcmd in tempcmdfinal:
                                        if tempcmd[-1] != ' ':
                                            cmdfinal.append(tempcmd+' '+cmd)
                                        else:
                                            cmdfinal.append(tempcmd+cmd)
                                else:
                                    cmdfinal = [cmd]
                        for item in cmdfinal:
                            new_commandlist.append(item)
                else:
                    new_commandlist.append(command)
            file2[title][2] = new_commandlist
        # ', ' to ','
        for item in file2:
            temp_list = []
            for temp in file2[item][2]:
                temp_list.append(temp.replace(', ',','))
            file2[item][2] = temp_list
        return file2
    def start_mode(self,mode):
        os.system(terminal_cmd)
        try:
            os.remove(python_path+"/res/testanswer.csv")
        except:
            pass
        if self.monitor_main_port_item==None or self.monitor_main_port_item=='':
            subinput2="USB-C"
        else:
            subinput2=self.monitor_main_port_item
        if self.monitor_second_port_item==None or self.monitor_second_port_item=='':
            subinput3='USB-C'
        else:
            subinput3=self.monitor_second_port_item
        if  self.monitor_third_port_item==None or self.monitor_third_port_item=='':
            subinput4='USB-C'
        else:
            subinput4=self.monitor_third_port_item
        # 跳過的題目
        if self.temp_skip_list==[]:
            skip2=[]
        else:
            skip2=self.temp_skip_list
        
        preset1=self.temp_preset1_list+self.temp_preset3_list
        window=self.temp_window_list
        subinput=self.temp_input_list
        sheet=self.sheet_choose
        i=self.start_task_num
        i2=self.end_task_num
        print('題目=',i,'題目＝',i2,'sheet=',sheet,'preset1=',preset1,'window=',window,'subinput=',subinput,'main_port=',subinput2,'second_port=',subinput3,'third_port=',subinput4,'skip=',skip2)
        sheetcanbecreate=0
        if internet_on()==True:
            auth_json_path =python_path+'/res/ddm-test-answer-2021-d7a5933c871b.json'
            gss_scopes = ['https://spreadsheets.google.com/feeds']
            #連線
            #credentials = ServiceAccountCredentials.from_json_keyfile_name(auth_json_path,gss_scopes)
            credentials = Credentials.from_service_account_file(auth_json_path,scopes=gss_scopes)
            gss_client = gspread.authorize(credentials)
            #開啟 Google Sheet 資料表
            spreadsheet_key = '1YXFqt-ZuYGPlWIzjEdGGzy3-Rcj6kUV3zfhi7tfWET0' 
            #建立工作表1
            #sheet = gss_client.open_by_key(spreadsheet_key).sheet1
            #自定義工作表名稱
            ddtimestart=f'{dt2.now()}'
            ddtimestart_sheet_title=f"{dt2.now().strftime('%Y%m%d%H%M%S')}"
            try:
                googlename=sheet
            except:
                googlename=ddtimestart_sheet_title
                sheet = gss_client.open_by_key(spreadsheet_key).add_worksheet(googlename, 500, 10, index=None)
                values =["steporder","document","example","terminal","result","Passfail","Comment"]
                sheet.insert_row(values, 1) #插入values到第1列
                sheetcanbecreate+=1
            if sheetcanbecreate==0:
                try:
                    sheet = gss_client.open_by_key(spreadsheet_key).add_worksheet(googlename, 500, 10, index=None)
                    values =["steporder","document","example","terminal","result","Passfail","Comment"]
                    sheet.insert_row(values, 1) #插入values到第1列
                except:
                    sheet = gss_client.open_by_key(spreadsheet_key).worksheet(googlename)
                    values =["steporder","document","example","terminal","result","Passfail","Comment"]
                    sheet.insert_row(values, len(sheet.get_all_values())+1) #插入values到第1列
            #google sheet
        #reset DDM
        file = Path(ddpm_ddm(cmd_logfile,self.command_start_choose))
        if file.exists():
            filetxt=Path(ddpm_ddm(cmd_logfile,self.command_start_choose))
            if filetxt.exists():
                os.remove(ddpm_ddm(cmd_logfile,self.command_start_choose))
            filetxt2=Path(cmd_dialogfile)
            if filetxt2.exists():
                os.remove(cmd_dialogfile)
        else:
            t_end = time.time() + 2
            while time.time() < t_end:
                if self.script_iwantto_stop==True:
                    return
                elif self.script_temp_iwantto_stop==True:
                    while self.script_temp_iwantto_stop==True:
                        QApplication.processEvents()
                QApplication.processEvents()
            # self.typing("cd /Applications/DDM")
            # self.typing("mkdir temp")
        #input source

        #不跑的題目
        if mode =='auto':
            f = open(python_path+'/res/skip.json')
            temp_data = json.load(f)
            skip = temp_data['skip']
        else:
            skip = []
        with open(python_path+"/res/testanswer.csv",'a') as fd:
            writer = csv.writer(fd)
            writer.writerow(['title','task','example','command','ANS','PassFail','Comment'])
        ddct=getdict(beta_name)
        ddct = self.reverse_dict(ddct)
        for item in ddct:
            print('-------------')
            print(item)
            for item in ddct[item][2]:
                print(item)
        if mode == 'auto':
            if i =='':
                i=1
            try:
                i=int(i)
            except:
                i=1
            if i2=='':
                i2=len(ddct)
            try:
                i2=int(i2)
            except:
                i2=len(ddct)
            
            if i2>=len(ddct):
                i2=len(ddct)
            if i<=1:
                i=1
        else:
            if i =='':
                i=1
            try:
                i=int(i)
                i2=int(i2)
            except:
                i=1
                i2=i
        backup_dict = {}
        internet_bool = True
        try:
            number = len(sheet.get_all_values())
        except:
            number = 0
        while i<=i2:
            ## . 先輸入題目和examples
            values = [i,ddct[i][0],ddct[i][1],'','','','']
            print(i,'題length = ',len(backup_dict))
            try:
                sheet.insert_row(values, number+1) 
                number += 1
                backup_dict[number] = values
            except Exception as e:
                internet_bool = False
                number+=1
                backup_dict[number] = values
            ##
            if self.script_iwantto_stop==True:
                return
            elif self.script_temp_iwantto_stop==True:
                while self.script_temp_iwantto_stop==True:
                    QApplication.processEvents()
            cmdlist=''
            cmdans=''
            dictlist=ddct[i][2]
            print('dictlist = ',dictlist)
            # print("題目＝",i,"幾個cmd＝")
            print(i)
            tasklist=ddct[i][0].split('\n')
            k=0
            skipcmd=0
            if mode =='auto':
                while k<len(skip2):
                    QApplication.processEvents()
                    if 'Debug' in skip2[k]:
                        if 'Debug' in tasklist[0]:
                            skipcmd+=1
                    if 'Volume' in skip2[k]:
                        if 'Volume' in tasklist[0]:
                            skipcmd+=1
                    if 'Uniformity Compensation' in skip2[k]:
                        if 'Uniformity Compensation' in tasklist[0]:
                            skipcmd+=1
                    if 'Rotation' in skip2[k]:
                        if 'Rotation' in tasklist[0]:
                            skipcmd+=1
                    if 'PIP' in skip2[k]:
                        if 'PIP' in tasklist[0]:
                            skipcmd+=1
                    if 'PBP(2window)' in skip2[k]:
                        if 'PBP - 2 Window' in tasklist[0]:
                            skipcmd+=1
                    if 'PBP(3or4window)' in skip2[k]:
                        if 'PBP - 3 Windows' in tasklist[0] or 'PBP - 4 Windows' in tasklist[0]:
                            skipcmd+=1
                    k+=1
            if skipcmd>=1:
                j=0
                while j<len(dictlist):
                    cmd=f'{dictlist[j]}'
                    cmdlist=cmdlist+cmd+'\n'
                    j+=1
                if len(dictlist)>0:
                    cmdans=cmdans+"這題被選為不跑"+"\n"
                try:
                    sheet.update_cell(len(sheet.get_all_values()), 4, cmdlist)
                    sheet.update_cell(len(sheet.get_all_values()), 5, cmdans)
                except Exception as e:
                    print('error=2305',e)
                    internet_bool = False
                backup_dict[number][3] = cmdlist
                backup_dict[number][4] = cmdans
            elif skipcmd<1:
                if self.script_iwantto_stop==True:
                    return
                elif self.script_temp_iwantto_stop==True:
                    while self.script_temp_iwantto_stop==True:
                        QApplication.processEvents()
                if len(dictlist)>0:
                    if i not in skip:
                        j=0
                        while j<len(dictlist):
                            cmd=f'{dictlist[j]}'
                            print('command=',cmd)
                            print("題目=",i,"指令=",j,"長度=",len(dictlist))
                            if input_filter in cmd and i<multiple_task:
                                cmdlist=cmdlist+cmd+'\n'
                                cmdans=self.runcmd(cmdans,cmd)
                                if cmdans == False:
                                    return 
                                getcmd = single_input_read
                                try:
                                    sheet.update_cell(len(sheet.get_all_values()), 4, cmdlist)
                                    sheet.update_cell(len(sheet.get_all_values()), 5, cmdans)
                                except Exception as e:
                                    internet_bool = False
                                cmdlist = cmdlist +getcmd+'\n'
                                cmdans=self.runcmd(cmdans,getcmd)
                                if cmdans == False:
                                    return 
                                try:
                                    sheet.update_cell(len(sheet.get_all_values()), 4, cmdlist)
                                    sheet.update_cell(len(sheet.get_all_values()), 5, cmdans)
                                except Exception as e:
                                    internet_bool = False
                            elif input_filter in cmd and i>multiple_task:
                                cmdlist=cmdlist+cmd+'\n'
                                cmdans=self.runcmd(cmdans,cmd)
                                if cmdans == False:
                                    return 
                                getcmd = multi_input_read
                                try:
                                    sheet.update_cell(len(sheet.get_all_values()), 4, cmdlist)
                                    sheet.update_cell(len(sheet.get_all_values()), 5, cmdans)
                                except Exception as e:
                                    internet_bool = False
                                cmdlist = cmdlist +getcmd+'\n'
                                cmdans=self.runcmd(cmdans,getcmd)
                                if cmdans == False:
                                    return 
                                try:
                                    sheet.update_cell(len(sheet.get_all_values()), 4, cmdlist)
                                    sheet.update_cell(len(sheet.get_all_values()), 5, cmdans)
                                except Exception as e:
                                    internet_bool = False
                            elif preset_filter in cmd and i<multiple_task:
                                cmdlist=cmdlist+cmd+'\n'
                                cmdans=self.runcmd(cmdans,cmd)
                                if cmdans == False:
                                    return 
                                getcmd = single_preset_read
                                try:
                                    sheet.update_cell(len(sheet.get_all_values()), 4, cmdlist)
                                    sheet.update_cell(len(sheet.get_all_values()), 5, cmdans)
                                except Exception as e:
                                    internet_bool = False
                                cmdlist = cmdlist +getcmd+'\n'
                                cmdans=self.runcmd(cmdans,getcmd)
                                if cmdans == False:
                                    return 
                                try:
                                    sheet.update_cell(len(sheet.get_all_values()), 4, cmdlist)
                                    sheet.update_cell(len(sheet.get_all_values()), 5, cmdans)
                                except Exception as e:
                                    internet_bool = False
                            elif preset_filter in cmd and i>multiple_task:
                                cmdlist=cmdlist+cmd+'\n'
                                cmdans=self.runcmd(cmdans,cmd)
                                if cmdans == False:
                                    return 
                                getcmd = multi_preset_read
                                try:
                                    sheet.update_cell(len(sheet.get_all_values()), 4, cmdlist)
                                    sheet.update_cell(len(sheet.get_all_values()), 5, cmdans)
                                except Exception as e:
                                    internet_bool = False
                                cmdlist = cmdlist +getcmd+'\n'
                                cmdans=self.runcmd(cmdans,getcmd)
                                if cmdans == False:
                                    return 
                                try:
                                    sheet.update_cell(len(sheet.get_all_values()), 4, cmdlist)
                                    sheet.update_cell(len(sheet.get_all_values()), 5, cmdans)
                                except Exception as e:
                                    internet_bool = False
                            else:
                                cmdlist=cmdlist+cmd+'\n'
                                cmdans=self.runcmd(cmdans,cmd)
                                if cmdans == False:
                                    return 
                                try:
                                    sheet.update_cell(len(sheet.get_all_values()), 4, cmdlist)
                                    sheet.update_cell(len(sheet.get_all_values()), 5, cmdans)
                                except Exception as e:
                                    internet_bool = False
                            backup_dict[number][3] = cmdlist
                            backup_dict[number][4] = cmdans
                            j+=1
                    else:
                        j=0
                        while j<len(dictlist):
                            cmd=f'{dictlist[j]}'
                            cmdlist=cmdlist+cmd+'\n'
                            j+=1
                        if len(dictlist)>0:
                            cmdans=cmdans+"這題被選為不跑"+"\n"
                        try:
                            sheet.update_cell(len(sheet.get_all_values()), 4, cmdlist)
                            sheet.update_cell(len(sheet.get_all_values()), 5, cmdans)
                        except Exception as e:
                            internet_bool = False
                        backup_dict[number][3] = cmdlist
                        backup_dict[number][4] = cmdans
                    if mode == 'auto':
                        if input_filter in cmdlist:
                            if '2:' in cmdlist:
                                self.typing(reset_input_2+' '+subinput3)
                            self.typing(reset_input_1+' '+subinput2)

                        if 'PIP' in ddct[i][0] or 'PBP' in ddct[i][0]:
                            self.typing(pip_off)
                        if 'Language' in ddct[i][0]:
                            self.typing(reset_language)
                    
            try:
                if len(cmdans)<1 or cmdans=='這題被選為不跑\n':
                    pass
                else:
                    t_end = time.time() + int(self.time_choose_num)
                    while time.time() < t_end:
                        if self.script_iwantto_stop==True:
                            return
                        elif self.script_temp_iwantto_stop==True:
                            while self.script_temp_iwantto_stop==True:
                                QApplication.processEvents()
                        QApplication.processEvents()
                print(cmdans)
            except Exception as e:
                print(e)
            try:
                with open(python_path+"/res/testanswer.csv",'a') as fd:
                    writer = csv.writer(fd)
                    writer.writerow([i,ddct[i][0],ddct[i][1],cmdlist,cmdans,'',''])
            except:
                pass
            i+=1
        #結束時輸出需要手動的題目
        if mode == 'auto':
            ddtime=f'{dt2.now()}'
            finalhand=f'{list(sorted(skip))}'
            try:
                with open(python_path+"/res/testanswer.csv",'a') as fd:
                    writer = csv.writer(fd)
                    writer.writerow([i,skip,'手動模式處理 ',ddtimestart,ddtime,'',''])
            except:
                pass    
            try:
                for number_temp in backup_dict:
                    print(number_temp,backup_dict[number_temp])
                    sheet.batch_update([{
                            'range': 'A'+str(number_temp)+':G'+str(number_temp),
                            'values': [backup_dict[number_temp]],
                        }])
            except Exception as e:
                print('error=',e)      
            while internet_bool == False:
                time.sleep(10)  
                if internet_on()==True:
                    internet_bool = True

            try:
                for number_temp in backup_dict:
                    print(number_temp,backup_dict[number_temp])
                    sheet.batch_update([{
                            'range': 'A'+str(number_temp)+':G'+str(number_temp),
                            'values': [backup_dict[number_temp]],
                        }])
            except Exception as e:
                print('error=',e)
            try:
                values = [i,finalhand,'手動模式處理 ',ddtimestart,ddtime,'','','']
                sheet.insert_row(values, number+1) 
            except Exception as e:
                print('error=2530',e)
            
    
        self.stop()
    # 單一程式
    
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = main()
    window.show()
    sys.exit(app.exec())
    
    
    
    
