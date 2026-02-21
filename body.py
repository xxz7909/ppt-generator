10_93# coding:utf-8
from selenium import webdriver
from PIL import Image
from PIL import ImageGrab
import time
import pyautogui
import pandas as pd
import xlwt
import sys
import os
import xlsxwriter
import xlrd
from xlrd import xldate_as_tuple
import datetime
import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email import encoders
from email.mime.base import MIMEBase
from email.utils import parseaddr, formataddr
import base64
from email.mime.image import MIMEImage
import traceback
from bypy import ByPy
import xlwings as xw
from time import sleep
from shutil import copyfile

#建立txt原始数据库
class Logger(object):
    def __init__(self, filename="Default.log"):
        self.terminal = sys.stdout
        self.log = open(filename, "a")

    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)

    def flush(self):
        pass

path = os.path.abspath(os.path.dirname(__file__))
type = sys.getfilesystemencoding()
sys.stdout = Logger('bumen.txt')

workbook = xlsxwriter.Workbook('C:/Users/Administrator/Desktop/发送区B/bumen/'+ time.strftime('%m.%d',time.localtime(time.time())) + '新媒体数据.xlsx')
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold':True})

Browser = webdriver.Firefox(executable_path = r'C:\Users\Administrator\Desktop\geckodriver.exe')
time.sleep(4)

SJ = "此次数据抓取时间为： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))
print (SJ)

geshi = workbook.add_format({
    'align':    'center',
    'valign':   'vcenter',
})

biaotougeshi = workbook.add_format({
    'bold':     True,
    'align':    'center',
    'valign':   'vcenter',
})


worksheet.set_column('A:A',25)
worksheet.merge_range('A1:D1',SJ,biaotougeshi)



#央视三农
id_1_1 = '18197287211512700'
url_1_1 = 'https://m.yangshipin.cn/user?cpid=' + id_1_1
Browser.get(url_1_1)
time.sleep(4)
fs_1_1 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_1_1.text[-1].strip() == '万':
    fensi_1_1 = int(float(fs_1_1.text[:-1])*10000)
else:
    fensi_1_1 = fs_1_1.text

bf_1_1 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_1_1.text[-1].strip() == '亿':
    bofang_1_1 = int(float(bf_1_1.text[:-1])*100000000)
else:
    bofang_1_1 = bf_1_1.text

sc_1_1 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_1_1.text[-1].strip() == '万':
    shangchuan_1_1 = int(float(sc_1_1.text[:-1])*10000)
else:
    shangchuan_1_1 = sc_1_1.text

name_1_1 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')

print('账号为：',name_1_1.text + " ；上传量：",shangchuan_1_1 , ' ；播放量：' , bofang_1_1 , " ；粉丝量：" ,fensi_1_1)

worksheet.write('A2', '账号名称',biaotougeshi)
worksheet.write('A3', name_1_1.text,geshi)
worksheet.write('B2', '上传量',biaotougeshi)
worksheet.write('B3', shangchuan_1_1,geshi)
worksheet.write('C2', '播放量',biaotougeshi)
worksheet.write('C3', bofang_1_1,geshi)
worksheet.write('D2', '粉丝量',biaotougeshi)
worksheet.write('D3', fensi_1_1,geshi)
worksheet.write('E3', '融媒体部',geshi)

#地球村日记
id_1_2 = '18197286372567289'
url_1_2 = 'https://m.yangshipin.cn/user?cpid=' + id_1_2
Browser.get(url_1_2)
time.sleep(4)
fs_1_2 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_1_2.text[-1].strip() == '万':
    fensi_1_2 = int(float(fs_1_2.text[:-1])*10000)
else:
    fensi_1_2 = fs_1_2.text

bf_1_2 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_1_2.text[-1].strip() == '万':
    bofang_1_2 = int(float(bf_1_2.text[:-1])*10000)
else:
    bofang_1_2 = bf_1_2.text

sc_1_2 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_1_2.text[-1].strip() == '万':
    shangchuan_1_2 = int(float(sc_1_2.text[:-1])*10000)
else:
    shangchuan_1_2 = sc_1_2.text

name_1_2 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
biaotou1_2 = "此次" + name_1_2.text + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou1_2)
print('账号为：',name_1_2.text + " ；上传量：",shangchuan_1_2 , ' ；播放量：' , bofang_1_2 , " ；粉丝量：" ,fensi_1_2)

worksheet.write('A4', name_1_2.text,geshi)
worksheet.write('B4', shangchuan_1_2,geshi)
worksheet.write('C4', bofang_1_2,geshi)
worksheet.write('D4', fensi_1_2,geshi)
worksheet.write('E4', '融媒体部',geshi)

#我爱发明
id_1_3 = '18197287043681913'
url_1_3 = 'https://m.yangshipin.cn/user?cpid=' + id_1_3
Browser.get(url_1_3)
time.sleep(4)
fs_1_3 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_1_3.text[-1].strip() == '万':
    fensi_1_3 = int(float(fs_1_3.text[:-1])*10000)
else:
    fensi_1_3 = fs_1_3.text

bf_1_3 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_1_3.text[-1].strip() == '亿':
    bofang_1_3 = int(float(bf_1_3.text[:-1])*100000000)
else:
    bofang_1_3 = bf_1_3.text

sc_1_3 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_1_3.text[-1].strip() == '万':
    shangchuan_1_3 = int(float(sc_1_3.text[:-1])*10000)
else:
    shangchuan_1_3 = sc_1_3.text

name_1_3 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
biaotou1_3 = "此次" + name_1_3.text + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou1_3)
print('账号为：',name_1_3.text + " ；上传量：",shangchuan_1_3 , ' ；播放量：' , bofang_1_3 , " ；粉丝量：" ,fensi_1_3)

worksheet.write('A5', name_1_3.text,geshi)
worksheet.write('B5', shangchuan_1_3,geshi)
worksheet.write('C5', bofang_1_3,geshi)
worksheet.write('D5', fensi_1_3,geshi)
worksheet.write('E5', '项目部',geshi)

#谁知盘中餐
id_1_4 = '18197286875883928'
url_1_4 = 'https://m.yangshipin.cn/user?cpid=' + id_1_4
Browser.get(url_1_4)
time.sleep(4)
fs_1_4 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_1_4.text[-1].strip() == '万':
    fensi_1_4 = int(float(fs_1_4.text[:-1])*10000)
else:
    fensi_1_4 = fs_1_4.text

bf_1_4 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_1_4.text[-1].strip() == '万':
    bofang_1_4 = int(float(bf_1_4.text[:-1])*10000)
else:
    bofang_1_4 = bf_1_4.text

sc_1_4 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_1_4.text[-1].strip() == '万':
    shangchuan_1_4 = int(float(sc_1_4.text[:-1])*10000)
else:
    shangchuan_1_4 = sc_1_4.text

name_1_4 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
biaotou1_4 = "此次" + name_1_4.text + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou1_4)
print('账号为：',name_1_4.text + " ；上传量：",shangchuan_1_4 , ' ；播放量：' , bofang_1_4 , " ；粉丝量：" ,fensi_1_4)

worksheet.write('A6', name_1_4.text,geshi)
worksheet.write('B6', shangchuan_1_4,geshi)
worksheet.write('C6', bofang_1_4,geshi)
worksheet.write('D6', fensi_1_4,geshi)
worksheet.write('E6', '合作节目部',geshi)

#我的家乡我代言
id_1_5 = '21907741167380533'
url_1_5 = 'https://m.yangshipin.cn/user?cpid=' + id_1_5
Browser.get(url_1_5)
time.sleep(4)
fs_1_5 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_1_5.text[-1].strip() == '万':
    fensi_1_5 = int(float(fs_1_5.text[:-1])*10000)
else:
    fensi_1_5 = fs_1_5.text

bf_1_5 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_1_5.text[-1].strip() == '万':
    bofang_1_5 = int(float(bf_1_5.text[:-1])*10000)
else:
    bofang_1_5 = bf_1_5.text

sc_1_5 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_1_5.text[-1].strip() == '万':
    shangchuan_1_5 = int(float(sc_1_5.text[:-1])*10000)
else:
    shangchuan_1_5 = sc_1_5.text

name_1_5 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
biaotou1_5 = "此次" + name_1_5.text + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou1_5)
print('账号为：',name_1_5.text + " ；上传量：",shangchuan_1_5 , ' ；播放量：' , bofang_1_5 , " ；粉丝量：" ,fensi_1_5)

worksheet.write('A7', name_1_5.text,geshi)
worksheet.write('B7', shangchuan_1_5,geshi)
worksheet.write('C7', bofang_1_5,geshi)
worksheet.write('D7', fensi_1_5,geshi)
worksheet.write('E7', '项目部',geshi)

#田园频道
id_1_6 = '21907803746481065'
url_1_6 = 'https://m.yangshipin.cn/user?cpid=' + id_1_6
Browser.get(url_1_6)
time.sleep(4)
fs_1_6 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_1_6.text[-1].strip() == '万':
    fensi_1_6 = int(float(fs_1_6.text[:-1])*10000)
else:
    fensi_1_6 = fs_1_6.text

bf_1_6 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_1_6.text[-1].strip() == '万':
    bofang_1_6 = int(float(bf_1_6.text[:-1])*10000)
else:
    bofang_1_6 = bf_1_6.text

sc_1_6 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_1_6.text[-1].strip() == '万':
    shangchuan_1_6 = int(float(sc_1_6.text[:-1])*10000)
else:
    shangchuan_1_6 = sc_1_6.text

name_1_6 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
biaotou1_6 = "此次" + name_1_6.text + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou1_6)
print('账号为：',name_1_6.text + " ；上传量：",shangchuan_1_6 , ' ；播放量：' , bofang_1_6 , " ；粉丝量：" ,fensi_1_6)

worksheet.write('A8', name_1_6.text,geshi)
worksheet.write('B8', shangchuan_1_6,geshi)
worksheet.write('C8', bofang_1_6,geshi)
worksheet.write('D8', fensi_1_6,geshi)
worksheet.write('E8', '项目部',geshi)

#家政那些事儿
id_1_8 = '18197287043656295'
url_1_8 = 'https://m.yangshipin.cn/user?cpid=' + id_1_8
Browser.get(url_1_8)
time.sleep(4)
fs_1_8 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_1_8.text[-1].strip() == '万':
    fensi_1_8 = int(float(fs_1_8.text[:-1])*10000)
else:
    fensi_1_8 = fs_1_8.text

bf_1_8 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_1_8.text[-1].strip() == '亿':
    bofang_1_8 = int(float(bf_1_8.text[:-1])*100000000)
else:
    bofang_1_8 = bf_1_8.text

sc_1_8 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_1_8.text[-1].strip() == '万':
    shangchuan_1_8 = int(float(sc_1_8.text[:-1])*10000)
else:
    shangchuan_1_8 = sc_1_8.text

name_1_8 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
biaotou1_8 = "此次" + name_1_8.text + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou1_8)
print('账号为：',name_1_8.text + " ；上传量：",shangchuan_1_8 , ' ；播放量：' , bofang_1_8 , " ；粉丝量：" ,fensi_1_8)

worksheet.write('A9', name_1_8.text,geshi)
worksheet.write('B9', shangchuan_1_8,geshi)
worksheet.write('C9', bofang_1_8,geshi)
worksheet.write('D9', fensi_1_8,geshi)
worksheet.write('E9', '项目部',geshi)

#医路有你
id_1_9 = '22819938869588992'
url_1_9 = 'https://m.yangshipin.cn/user?cpid=' + id_1_9
Browser.get(url_1_9)
time.sleep(4)
fs_1_9 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_1_9.text[-1].strip() == '万':
    fensi_1_9 = int(float(fs_1_9.text[:-1])*10000)
else:
    fensi_1_9 = fs_1_9.text

bf_1_9 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_1_9.text[-1].strip() == '万':
    bofang_1_9 = int(float(bf_1_9.text[:-1])*10000)
else:
    bofang_1_9 = bf_1_9.text

sc_1_9 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_1_9.text[-1].strip() == '万':
    shangchuan_1_9 = int(float(sc_1_9.text[:-1])*10000)
else:
    shangchuan_1_9 = sc_1_9.text

name_1_9 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
biaotou1_9 = "此次" + name_1_9.text + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou1_9)
print('账号为：',name_1_9.text + " ；上传量：",shangchuan_1_9 , ' ；播放量：' , bofang_1_9 , " ；粉丝量：" ,fensi_1_9)

worksheet.write('A10', name_1_9.text,geshi)
worksheet.write('B10', shangchuan_1_9,geshi)
worksheet.write('C10', bofang_1_9,geshi)
worksheet.write('D10', fensi_1_9,geshi)
worksheet.write('E10', '项目部',geshi)

#乡村振兴观察点（央视频）
id_2_1 = '18197286540459682'
url_2_1 = 'https://m.yangshipin.cn/user?cpid=' + id_2_1
Browser.get(url_2_1)
time.sleep(4)
fs_2_1 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_2_1.text[-1].strip() == '万':
    fensi_2_1 = int(float(fs_2_1.text[:-1])*10000)
else:
    fensi_2_1 = fs_2_1.text

bf_2_1 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_2_1.text[-1].strip() == '万':
    bofang_2_1 = int(float(bf_2_1.text[:-1])*10000)
else:
    bofang_2_1 = bf_2_1.text

sc_2_1 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_2_1.text[-1].strip() == '万':
    shangchuan_2_1 = int(float(sc_2_1.text[:-1])*10000)
else:
    shangchuan_2_1 = sc_2_1.text

name_2_1_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
name_2_1 = name_2_1_0.text + '(央视频)'
biaotou2_1 = "此次" + name_2_1 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou2_1)
print('账号为：',name_2_1 + " ；上传量：",shangchuan_2_1 , ' ；播放量：' , bofang_2_1 , " ；粉丝量：" ,fensi_2_1)

worksheet.set_column('H:H',20)

worksheet.write('A11', name_2_1,geshi)
worksheet.write('B11', shangchuan_2_1,geshi)
worksheet.write('C11', bofang_2_1,geshi)
worksheet.write('D11', fensi_2_1,geshi)
worksheet.write('E11', '合作节目部',geshi)

#共富经
id_2_2 = '18197287379308107'
url_2_2 = 'https://m.yangshipin.cn/user?cpid=' + id_2_2
Browser.get(url_2_2)
time.sleep(4)
fs_2_2 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_2_2.text[-1].strip() == '万':
    fensi_2_2 = int(float(fs_2_2.text[:-1])*10000)
else:
    fensi_2_2 = fs_2_2.text

bf_2_2 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_2_2.text[-1].strip() == '万':
    bofang_2_2 = int(float(bf_2_2.text[:-1])*10000)
else:
    bofang_2_2 = bf_2_2.text

sc_2_2 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_2_2.text[-1].strip() == '万':
    shangchuan_2_2 = int(float(sc_2_2.text[:-1])*10000)
else:
    shangchuan_2_2 = sc_2_2.text

name_2_2_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
name_2_2 = name_2_2_0.text + '(央视频)'
biaotou2_2 = "此次" + name_2_2 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou2_2)
print('账号为：',name_2_2 + " ；上传量：",shangchuan_2_2 , ' ；播放量：' , bofang_2_2 , " ；粉丝量：" ,fensi_2_2)

worksheet.write('A12', name_2_2,geshi)
worksheet.write('B12', shangchuan_2_2,geshi)
worksheet.write('C12', bofang_2_2,geshi)
worksheet.write('D12', fensi_2_2,geshi)
worksheet.write('E12', '合作节目部',geshi)

#田园帮帮团(央视频)
id_2_3 = '18197286875991119'
url_2_3 = 'https://m.yangshipin.cn/user?cpid=' + id_2_3
Browser.get(url_2_3)
time.sleep(4)
fs_2_3 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_2_3.text[-1].strip() == '万':
    fensi_2_3 = int(float(fs_2_3.text[:-1])*10000)
else:
    fensi_2_3 = fs_2_3.text

bf_2_3 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_2_3.text[-1].strip() == '万':
    bofang_2_3 = int(float(bf_2_3.text[:-1])*10000)
else:
    bofang_2_3 = bf_2_3.text

sc_2_3 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_2_3.text[-1].strip() == '万':
    shangchuan_2_3 = int(float(sc_2_3.text[:-1])*10000)
else:
    shangchuan_2_3 = sc_2_3.text

name_2_3_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
name_2_3 = name_2_3_0.text + '(央视频)'
biaotou2_3 = "此次" + name_2_3 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou2_3)
print('账号为：',name_2_3 + " ；上传量：",shangchuan_2_3 , ' ；播放量：' , bofang_2_3 , " ；粉丝量：" ,fensi_2_3)

worksheet.write('A13', name_2_3,geshi)
worksheet.write('B13', shangchuan_2_3,geshi)
worksheet.write('C13', bofang_2_3,geshi)
worksheet.write('D13', fensi_2_3,geshi)
worksheet.write('E13', '合作节目部',geshi)

#央视主播说三农（央视频）
id_0_1 = '18197286372644427'
url_0_1 = 'https://m.yangshipin.cn/user?cpid=' + id_0_1
Browser.get(url_0_1)
time.sleep(4)
fs_0_1 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_0_1.text[-1].strip() == '万':
    fensi_0_1 = int(float(fs_0_1.text[:-1])*10000)
else:
    fensi_0_1 = fs_0_1.text

bf_0_1 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_0_1.text[-1].strip() == '万':
    bofang_0_1 = int(float(bf_0_1.text[:-1])*10000)
else:
    bofang_0_1 = bf_0_1.text

sc_0_1 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_0_1.text[-1].strip() == '万':
    shangchuan_0_1 = int(float(sc_0_1.text[:-1])*10000)
else:
    shangchuan_0_1 = sc_0_1.text

name_0_1_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
name_0_1 = name_0_1_0.text + '(央视频)'
biaotou0_1 = "此次" + name_0_1 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou0_1)
print('账号为：',name_0_1 + " ；上传量：",shangchuan_0_1 , ' ；播放量：' , bofang_0_1 , " ；粉丝量：" ,fensi_0_1)

worksheet.write('A14', name_0_1,geshi)
worksheet.write('B14', shangchuan_0_1,geshi)
worksheet.write('C14', bofang_0_1,geshi)
worksheet.write('D14', fensi_0_1,geshi)
worksheet.write('E14', '新闻部',geshi)

#大地讲堂
id_3_1 = '18197286372687491'
url_3_1 = 'https://m.yangshipin.cn/user?cpid=' + id_3_1
Browser.get(url_3_1)
time.sleep(4)
fs_3_1 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_3_1.text[-1].strip() == '万':
    fensi_3_1 = int(float(fs_3_1.text[:-1])*10000)
else:
    fensi_3_1 = fs_3_1.text

bf_3_1 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_3_1.text[-1].strip() == '万':
    bofang_3_1 = int(float(bf_3_1.text[:-1])*10000)
else:
    bofang_3_1 = bf_3_1.text

sc_3_1 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_3_1.text[-1].strip() == '万':
    shangchuan_3_1 = int(float(sc_3_1.text[:-1])*10000)
else:
    shangchuan_3_1 = sc_3_1.text

name_3_1_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
name_3_1 = name_3_1_0.text + '(央视频)'
biaotou3_1 = "此次" + name_3_1 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou3_1)
print('账号为：',name_3_1 + " ；上传量：",shangchuan_3_1 , ' ；播放量：' , bofang_3_1 , " ；粉丝量：" ,fensi_3_1)

worksheet.write('A15', name_3_1,geshi)
worksheet.write('B15', shangchuan_3_1,geshi)
worksheet.write('C15', bofang_3_1,geshi)
worksheet.write('D15', fensi_3_1,geshi)
worksheet.write('E15', '专题节目部',geshi)

#乡理乡亲
id_3_2 = '18197287043740297'
url_3_2 = 'https://m.yangshipin.cn/user?cpid=' + id_3_2
Browser.get(url_3_2)
time.sleep(4)
fs_3_2 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_3_2.text[-1].strip() == '万':
    fensi_3_2 = int(float(fs_3_2.text[:-1])*10000)
else:
    fensi_3_2 = fs_3_2.text

bf_3_2 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_3_2.text[-1].strip() == '万':
    bofang_3_2 = int(float(bf_3_2.text[:-1])*10000)
else:
    bofang_3_2 = bf_3_2.text

sc_3_2 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_3_2.text[-1].strip() == '万':
    shangchuan_3_2 = int(float(sc_3_2.text[:-1])*10000)
else:
    shangchuan_3_2 = sc_3_2.text

name_3_2_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
name_3_2 = name_3_2_0.text + '(央视频)'
biaotou3_2 = "此次" + name_3_2 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou3_2)
print('账号为：',name_3_2+ " ；上传量：",shangchuan_3_2 , ' ；播放量：' , bofang_3_2 , " ；粉丝量：" ,fensi_3_2)

worksheet.write('A16', name_3_2,geshi)
worksheet.write('B16', shangchuan_3_2,geshi)
worksheet.write('C16', bofang_3_2,geshi)
worksheet.write('D16', fensi_3_2,geshi)
worksheet.write('E16', '专题节目部',geshi)

#乡土中国
id_3_3 = '18197286708219316'
url_3_3 = 'https://m.yangshipin.cn/user?cpid=' + id_3_3
Browser.get(url_3_3)
time.sleep(4)
fs_3_3 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_3_3.text[-1].strip() == '万':
    fensi_3_3 = int(float(fs_3_3.text[:-1])*10000)
else:
    fensi_3_3 = fs_3_3.text

bf_3_3 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_3_3.text[-1].strip() == '万':
    bofang_3_3 = int(float(bf_3_3.text[:-1])*10000)
else:
    bofang_3_3 = bf_3_3.text

sc_3_3 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_3_3.text[-1].strip() == '万':
    shangchuan_3_3 = int(float(sc_3_3.text[:-1])*10000)
else:
    shangchuan_3_3 = sc_3_3.text

name_3_3 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
biaotou3_3 = "此次" + name_3_3.text + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou3_3)
print('账号为：',name_3_3.text + " ；上传量：",shangchuan_3_3 , ' ；播放量：' , bofang_3_3 , " ；粉丝量：" ,fensi_3_3)

worksheet.write('A17', name_3_3.text,geshi)
worksheet.write('B17', shangchuan_3_3,geshi)
worksheet.write('C17', bofang_3_3,geshi)
worksheet.write('D17', fensi_3_3,geshi)
worksheet.write('E17', '专题节目部',geshi)

#小满GO乡村
id_3_4 = '18197286372593228'
url_3_4 = 'https://m.yangshipin.cn/user?cpid=' + id_3_4
Browser.get(url_3_4)
time.sleep(4)
fs_3_4 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_3_4.text[-1].strip() == '万':
    fensi_3_4 = int(float(fs_3_4.text[:-1])*10000)
else:
    fensi_3_4 = fs_3_4.text

bf_3_4 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_3_4.text[-1].strip() == '万':
    bofang_3_4 = int(float(bf_3_4.text[:-1])*10000)
else:
    bofang_3_4 = bf_3_4.text

sc_3_4 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_3_4.text[-1].strip() == '万':
    shangchuan_3_4 = int(float(sc_3_4.text[:-1])*10000)
else:
    shangchuan_3_4 = sc_3_4.text

name_3_4 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
biaotou3_4 = "此次" + name_3_4.text + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou3_4)
print('账号为：',name_3_4.text + " ；上传量：",shangchuan_3_4 , ' ；播放量：' , bofang_3_4 , " ；粉丝量：" ,fensi_3_4)

worksheet.write('A18', name_3_4.text,geshi)
worksheet.write('B18', shangchuan_3_4,geshi)
worksheet.write('C18', bofang_3_4,geshi)
worksheet.write('D18', fensi_3_4,geshi)
worksheet.write('E18', '项目部',geshi)

#三农群英汇
id_3_5 = '18197286708137450'
url_3_5 = 'https://m.yangshipin.cn/user?cpid=' + id_3_5
Browser.get(url_3_5)
time.sleep(4)
fs_3_5 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_3_5.text[-1].strip() == '万':
    fensi_3_5 = int(float(fs_3_5.text[:-1])*10000)
else:
    fensi_3_5 = fs_3_5.text

bf_3_5 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_3_5.text[-1].strip() == '万':
    bofang_3_5 = int(float(bf_3_5.text[:-1])*10000)
else:
    bofang_3_5 = bf_3_5.text

sc_3_5 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_3_5.text[-1].strip() == '万':
    shangchuan_3_5 = int(float(sc_3_5.text[:-1])*10000)
else:
    shangchuan_3_5 = sc_3_5.text

name_3_5 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
biaotou3_5 = "此次" + name_3_5.text + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou3_5)
print('账号为：',name_3_5.text + " ；上传量：",shangchuan_3_5 , ' ；播放量：' , bofang_3_5 , " ；粉丝量：" ,fensi_3_5)

worksheet.write('A19', name_3_5.text,geshi)
worksheet.write('B19', shangchuan_3_5,geshi)
worksheet.write('C19', bofang_3_5,geshi)
worksheet.write('D19', fensi_3_5,geshi)
worksheet.write('E19', '专题节目部',geshi)

#动物萌友会
id_3_6 = '19844777610398987'
url_3_6 = 'https://m.yangshipin.cn/user?cpid=' + id_3_6
Browser.get(url_3_6)
time.sleep(4)
fs_3_6 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_3_6.text[-1].strip() == '万':
    fensi_3_6 = int(float(fs_3_6.text[:-1])*10000)
else:
    fensi_3_6 = fs_3_6.text

bf_3_6 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_3_6.text[-1].strip() == '万':
    bofang_3_6 = int(float(bf_3_6.text[:-1])*10000)
else:
    bofang_3_6 = bf_3_6.text

sc_3_6 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_3_6.text[-1].strip() == '万':
    shangchuan_3_6 = int(float(sc_3_6.text[:-1])*10000)
else:
    shangchuan_3_6 = sc_3_6.text

name_3_6 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
biaotou3_6 = "此次" + name_3_6.text + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou3_6)
print('账号为：',name_3_6.text + " ；上传量：",shangchuan_3_6 , ' ；播放量：' , bofang_3_6 , " ；粉丝量：" ,fensi_3_6)

worksheet.write('A20', name_3_6.text,geshi)
worksheet.write('B20', shangchuan_3_6,geshi)
worksheet.write('C20', bofang_3_6,geshi)
worksheet.write('D20', fensi_3_6,geshi)
worksheet.write('E20', '专题节目部',geshi)

#遇见民宿
id_6_1 = '22878853405665873'
url_6_1 = 'https://m.yangshipin.cn/user?cpid=' + id_6_1
Browser.get(url_6_1)
time.sleep(4)
fs_6_1 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_6_1.text[-1].strip() == '万':
    fensi_6_1 = int(float(fs_6_1.text[:-1])*10000)
else:
    fensi_6_1 = fs_6_1.text

bf_6_1 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_6_1.text[-1].strip() == '万':
    bofang_6_1 = int(float(bf_6_1.text[:-1])*10000)
else:
    bofang_6_1 = bf_6_1.text

sc_6_1 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_6_1.text[-1].strip() == '万':
    shangchuan_6_1 = int(float(sc_6_1.text[:-1])*10000)
else:
    shangchuan_6_1 = sc_6_1.text

name_6_1 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
biaotou6_1 = "此次" + name_6_1.text + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou6_1)
print('账号为：',name_6_1.text + " ；上传量：",shangchuan_6_1 , ' ；播放量：' , bofang_6_1 , " ；粉丝量：" ,fensi_6_1)

worksheet.write('A21', name_6_1.text,geshi)
worksheet.write('B21', shangchuan_6_1,geshi)
worksheet.write('C21', bofang_6_1,geshi)
worksheet.write('D21', fensi_6_1,geshi)
worksheet.write('E21', '项目部',geshi)

#玩转园艺
id_6_2 = '22878869176333111'
url_6_2 = 'https://m.yangshipin.cn/user?cpid=' + id_6_2
Browser.get(url_6_2)
time.sleep(4)
fs_6_2 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_6_2.text[-1].strip() == '万':
    fensi_6_2 = int(float(fs_6_2.text[:-1])*10000)
else:
    fensi_6_2 = fs_6_2.text

bf_6_2 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_6_2.text[-1].strip() == '万':
    bofang_6_2 = int(float(bf_6_2.text[:-1])*10000)
else:
    bofang_6_2 = bf_6_2.text

sc_6_2 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_6_2.text[-1].strip() == '万':
    shangchuan_6_2 = int(float(sc_6_2.text[:-1])*10000)
else:
    shangchuan_6_2 = sc_6_2.text

name_6_2 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
biaotou6_2 = "此次" + name_6_2.text + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou6_2)
print('账号为：',name_6_2.text + " ；上传量：",shangchuan_6_2 , ' ；播放量：' , bofang_6_2 , " ；粉丝量：" ,fensi_6_2)

worksheet.write('A22', name_6_2.text,geshi)
worksheet.write('B22', shangchuan_6_2,geshi)
worksheet.write('C22', bofang_6_2,geshi)
worksheet.write('D22', fensi_6_2,geshi)
worksheet.write('E22', '项目部',geshi)

#田野欢歌
id_6_3 = '18197287043776259'
url_6_3 = 'https://m.yangshipin.cn/user?cpid=' + id_6_3
Browser.get(url_6_3)
time.sleep(4)
fs_6_3 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_6_3.text[-1].strip() == '万':
    fensi_6_3 = int(float(fs_6_3.text[:-1])*10000)
else:
    fensi_6_3 = fs_6_3.text

bf_6_3 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_6_3.text[-1].strip() == '万':
    bofang_6_3 = int(float(bf_6_3.text[:-1])*10000)
else:
    bofang_6_3 = bf_6_3.text

sc_6_3 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_6_3.text[-1].strip() == '万':
    shangchuan_6_3 = int(float(sc_6_3.text[:-1])*10000)
else:
    shangchuan_6_3 = sc_6_3.text

name_6_3 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
biaotou6_3 = "此次" + name_6_3.text + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou6_3)
print('账号为：',name_6_3.text + " ；上传量：",shangchuan_6_3 , ' ；播放量：' , bofang_6_3 , " ；粉丝量：" ,fensi_6_3)

worksheet.write('A23', name_6_3.text,geshi)
worksheet.write('B23', shangchuan_6_3,geshi)
worksheet.write('C23', bofang_6_3,geshi)
worksheet.write('D23', fensi_6_3,geshi)
worksheet.write('E23', '文艺节目部',geshi)

#央视乡村文艺
id_6_4 = '22906403943173455'
url_6_4 = 'https://m.yangshipin.cn/user?cpid=' + id_6_4
Browser.get(url_6_4)
time.sleep(4)
fs_6_4 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_6_4.text[-1].strip() == '万':
    fensi_6_4 = int(float(fs_6_4.text[:-1])*10000)
else:
    fensi_6_4 = fs_6_4.text

bf_6_4 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_6_4.text[-1].strip() == '万':
    bofang_6_4 = int(float(bf_6_4.text[:-1])*10000)
else:
    bofang_6_4 = bf_6_4.text

sc_6_4 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_6_4.text[-1].strip() == '万':
    shangchuan_6_4 = int(float(sc_6_4.text[:-1])*10000)
else:
    shangchuan_6_4 = sc_6_4.text

name_6_4 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
biaotou6_4 = "此次" + name_6_4.text + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou6_4)
print('账号为：',name_6_4.text + " ；上传量：",shangchuan_6_4 , ' ；播放量：' , bofang_6_4 , " ；粉丝量：" ,fensi_6_4)

worksheet.write('A24', name_6_4.text,geshi)
worksheet.write('B24', shangchuan_6_4,geshi)
worksheet.write('C24', bofang_6_4,geshi)
worksheet.write('D24', fensi_6_4,geshi)
worksheet.write('E24', '文艺节目部',geshi)

#乐游新乡村
id_6_5 = '18197286708195953'
url_6_5 = 'https://m.yangshipin.cn/user?cpid=' + id_6_5
Browser.get(url_6_5)
time.sleep(4)
fs6_5 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs6_5.text[-1].strip() == '万':
    fensi6_5 = int(float(fs6_5.text[:-1])*10000)
else:
    fensi6_5 = fs6_5.text

bf6_5 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf6_5.text[-1].strip() == '万':
    bofang6_5 = int(float(bf6_5.text[:-1])*10000)
else:
    bofang6_5 = bf6_5.text

sc6_5 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc6_5.text[-1].strip() == '万':
    shangchuan6_5 = int(float(sc6_5.text[:-1])*10000)
else:
    shangchuan6_5 = sc6_5.text

name6_5_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
name6_5 = name6_5_0.text + '(央视频)'
biaotou6_5 = "此次" + name6_5 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou6_5)
print('账号为：',name6_5 + " ；上传量：",shangchuan6_5 , ' ；播放量：' , bofang6_5 , " ；粉丝量：" ,fensi6_5)

worksheet.write('A25', name6_5,geshi)
worksheet.write('B25', shangchuan6_5,geshi)
worksheet.write('C25', bofang6_5,geshi)
worksheet.write('D25', fensi6_5,geshi)
worksheet.write('E25', '文艺节目部',geshi)

#田野麦霸
id_6_6 = '23195627879721410'
url_6_6 = 'https://m.yangshipin.cn/user?cpid=' + id_6_6
Browser.get(url_6_6)
time.sleep(4)
fs_6_6 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_6_6.text[-1].strip() == '万':
    fensi_6_6 = int(float(fs_6_6.text[:-1])*10000)
else:
    fensi_6_6 = fs_6_6.text

bf_6_6 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_6_6.text[-1].strip() == '亿':
    bofang_6_6 = int(float(bf_6_6.text[:-1])*100000000)
else:
    bofang_6_6 = bf_6_6.text

sc_6_6 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_6_6.text[-1].strip() == '万':
    shangchuan_6_6 = int(float(sc_6_6.text[:-1])*10000)
else:
    shangchuan_6_6 = sc_6_6.text

name_6_6 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
biaotou6_6 = "此次" + name_6_6.text + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou6_6)
print('账号为：',name_6_6.text + " ；上传量：",shangchuan_6_6 , ' ；播放量：' , bofang_6_6 , " ；粉丝量：" ,fensi_6_6)

worksheet.write('A26', name_6_6.text,geshi)
worksheet.write('B26', shangchuan_6_6,geshi)
worksheet.write('C26', bofang_6_6,geshi)
worksheet.write('D26', fensi_6_6,geshi)
worksheet.write('E26', '项目部',geshi)

#乡村振兴面对面
id_6_7 = '18197287211535701'
url_6_7 = 'https://m.yangshipin.cn/user?cpid=' + id_6_7
Browser.get(url_6_7)
time.sleep(4)
fs_6_7 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_6_7.text[-1].strip() == '万':
    fensi_6_7 = int(float(fs_6_7.text[:-1])*10000)
else:
    fensi_6_7 = fs_6_7.text

bf_6_7 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_6_7.text[-1].strip() == '万':
    bofang_6_7 = int(float(bf_6_7.text[:-1])*10000)
else:
    bofang_6_7 = bf_6_7.text

sc_6_7 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_6_7.text[-1].strip() == '万':
    shangchuan_6_7 = int(float(sc_6_7.text[:-1])*10000)
else:
    shangchuan_6_7 = sc_6_7.text

name_6_7 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
biaotou6_7 = "此次" + name_6_7.text + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou6_7)
print('账号为：',name_6_7.text + " ；上传量：",shangchuan_6_7 , ' ；播放量：' , bofang_6_7 , " ；粉丝量：" ,fensi_6_7)

worksheet.write('A27', name_6_7.text,geshi)
worksheet.write('B27', shangchuan_6_7,geshi)
worksheet.write('C27', bofang_6_7,geshi)
worksheet.write('D27', fensi_6_7,geshi)
worksheet.write('E27', '无部门',geshi)

#乡约
id_6_8 = '18197287043760733'
url_6_8 = 'https://m.yangshipin.cn/user?cpid=' + id_6_8
Browser.get(url_6_8)
time.sleep(4)
fs_6_8 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_6_8.text[-1].strip() == '万':
    fensi_6_8 = int(float(fs_6_8.text[:-1])*10000)
else:
    fensi_6_8 = fs_6_8.text

bf_6_8 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_6_8.text[-1].strip() == '万':
    bofang_6_8 = int(float(bf_6_8.text[:-1])*10000)
else:
    bofang_6_8 = bf_6_8.text

sc_6_8 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_6_8.text[-1].strip() == '万':
    shangchuan_6_8 = int(float(sc_6_8.text[:-1])*10000)
else:
    shangchuan_6_8 = sc_6_8.text

name_6_8 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
biaotou6_8 = "此次" + name_6_8.text + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou6_8)
print('账号为：',name_6_8.text + " ；上传量：",shangchuan_6_8 , ' ；播放量：' , bofang_6_8 , " ；粉丝量：" ,fensi_6_8)

worksheet.write('A28', name_6_8.text,geshi)
worksheet.write('B28', shangchuan_6_8,geshi)
worksheet.write('C28', bofang_6_8,geshi)
worksheet.write('D28', fensi_6_8,geshi)
worksheet.write('E28', '文艺节目部',geshi)

#醒醒别上当
id_6_9 = '23427787639311702'
url_6_9 = 'https://m.yangshipin.cn/user?cpid=' + id_6_9
Browser.get(url_6_9)
time.sleep(4)
fs_6_9 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_6_9.text[-1].strip() == '万':
    fensi_6_9 = int(float(fs_6_9.text[:-1])*10000)
else:
    fensi_6_9 = fs_6_9.text

bf_6_9 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_6_9.text[-1].strip() == '亿':
    bofang_6_9 = int(float(bf_6_9.text[:-1])*100000000)
else:
    bofang_6_9 = bf_6_9.text

sc_6_9 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_6_9.text[-1].strip() == '万':
    shangchuan_6_9 = int(float(sc_6_9.text[:-1])*10000)
else:
    shangchuan_6_9 = sc_6_9.text

name_6_9 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
biaotou6_9 = "此次" + name_6_9.text + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou6_9)
print('账号为：',name_6_9.text + " ；上传量：",shangchuan_6_9 , ' ；播放量：' , bofang_6_9 , " ；粉丝量：" ,fensi_6_9)

worksheet.write('A29', name_6_9.text,geshi)
worksheet.write('B29', shangchuan_6_9,geshi)
worksheet.write('C29', bofang_6_9,geshi)
worksheet.write('D29', fensi_6_9,geshi)
worksheet.write('E29', '项目部',geshi)

#中国好时节
id_6_10 = '23442649232763161'
url_6_10 = 'https://m.yangshipin.cn/user?cpid=' + id_6_10
Browser.get(url_6_10)
time.sleep(4)
fs_6_10 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_6_10.text[-1].strip() == '万':
    fensi_6_10 = int(float(fs_6_10.text[:-1])*10000)
else:
    fensi_6_10 = fs_6_10.text

bf_6_10 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_6_10.text[-1].strip() == '万':
    bofang_6_10 = int(float(bf_6_10.text[:-1])*10000)
else:
    bofang_6_10 = bf_6_10.text

sc_6_10 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_6_10.text[-1].strip() == '万':
    shangchuan_6_10 = int(float(sc_6_10.text[:-1])*10000)
else:
    shangchuan_6_10 = sc_6_10.text

name_6_10 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
biaotou6_10 = "此次" + name_6_10.text + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou6_10)
print('账号为：',name_6_10.text + " ；上传量：",shangchuan_6_10 , ' ；播放量：' , bofang_6_10 , " ；粉丝量：" ,fensi_6_10)

worksheet.write('A30', name_6_10.text,geshi)
worksheet.write('B30', shangchuan_6_10,geshi)
worksheet.write('C30', bofang_6_10,geshi)
worksheet.write('D30', fensi_6_10,geshi)
worksheet.write('E30', '新闻部',geshi)

#乡村服务社
id_6_11 = '25428964065208652'
url_6_11 = 'https://m.yangshipin.cn/user?cpid=' + id_6_11
Browser.get(url_6_11)
time.sleep(4)
fs_6_11 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_6_11.text[-1].strip() == '万':
    fensi_6_11 = int(float(fs_6_11.text[:-1])*10000)
else:
    fensi_6_11 = fs_6_11.text

bf_6_11 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_6_11.text[-1].strip() == '万':
    bofang_6_11 = int(float(bf_6_11.text[:-1])*10000)
else:
    bofang_6_11 = bf_6_11.text

sc_6_11 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_6_11.text[-1].strip() == '万':
    shangchuan_6_11 = int(float(sc_6_11.text[:-1])*10000)
else:
    shangchuan_6_11 = sc_6_11.text

name_6_11 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
biaotou6_11 = "此次" + name_6_11.text + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou6_11)
print('账号为：',name_6_11.text + " ；上传量：",shangchuan_6_11 , ' ；播放量：' , bofang_6_11 , " ；粉丝量：" ,fensi_6_11)

worksheet.write('A31', name_6_11.text,geshi)
worksheet.write('B31', shangchuan_6_11,geshi)
worksheet.write('C31', bofang_6_11,geshi)
worksheet.write('D31', fensi_6_11,geshi)
worksheet.write('E31', '项目部',geshi)

#鲜知先觉
id_6_12 = '21907772708631613'
url_6_12 = 'https://m.yangshipin.cn/user?cpid=' + id_6_12
Browser.get(url_6_12)
time.sleep(4)
fs_6_12 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_6_12.text[-1].strip() == '万':
    fensi_6_12 = int(float(fs_6_12.text[:-1])*10000)
else:
    fensi_6_12 = fs_6_12.text

bf_6_12 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_6_12.text[-1].strip() == '亿':
    bofang_6_12 = int(float(bf_6_12.text[:-1])*100000000)
else:
    bofang_6_12 = bf_6_12.text

sc_6_12 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_6_12.text[-1].strip() == '万':
    shangchuan_6_12 = int(float(sc_6_12.text[:-1])*10000)
else:
    shangchuan_6_12 = sc_6_12.text

name_6_12 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
biaotou6_12 = "此次" + name_6_12.text + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou6_12)
print('账号为：',name_6_12.text + " ；上传量：",shangchuan_6_12 , ' ；播放量：' , bofang_6_12 , " ；粉丝量：" ,fensi_6_12)

worksheet.write('A32', name_6_12.text,geshi)
worksheet.write('B32', shangchuan_6_12,geshi)
worksheet.write('C32', bofang_6_12,geshi)
worksheet.write('D32', fensi_6_12,geshi)
worksheet.write('E32', '项目部',geshi)

#17朋友圈
id_6_13 = '27689907159649513'
url_6_13 = 'https://m.yangshipin.cn/user?cpid=' + id_6_13
Browser.get(url_6_13)
time.sleep(4)
fs_6_13 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_6_13.text[-1].strip() == '万':
    fensi_6_13 = int(float(fs_6_13.text[:-1])*10000)
else:
    fensi_6_13 = fs_6_13.text

bf_6_13 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_6_13.text[-1].strip() == '万':
    bofang_6_13 = int(float(bf_6_13.text[:-1])*10000)
else:
    bofang_6_13 = bf_6_13.text

sc_6_13 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_6_13.text[-1].strip() == '万':
    shangchuan_6_13 = int(float(sc_6_13.text[:-1])*10000)
else:
    shangchuan_6_13 = sc_6_13.text

name_6_13 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
biaotou6_13 = "此次" + name_6_13.text + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou6_13)
print('账号为：',name_6_13.text + " ；上传量：",shangchuan_6_13 , ' ；播放量：' , bofang_6_13 , " ；粉丝量：" ,fensi_6_13)

worksheet.write('A33', name_6_13.text,geshi)
worksheet.write('B33', shangchuan_6_13,geshi)
worksheet.write('C33', bofang_6_13,geshi)
worksheet.write('D33', fensi_6_13,geshi)
worksheet.write('E33', '广播节目部',geshi)

#振兴路上新观察
id_6_14 = '28211680758262146'
url_6_14 = 'https://m.yangshipin.cn/user?cpid=' + id_6_14
Browser.get(url_6_14)
time.sleep(4)
fs_6_14 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_6_14.text[-1].strip() == '万':
    fensi_6_14 = int(float(fs_6_14.text[:-1])*10000)
else:
    fensi_6_14 = fs_6_14.text

bf_6_14 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_6_14.text[-1].strip() == '万':
    bofang_6_14 = int(float(bf_6_14.text[:-1])*10000)
else:
    bofang_6_14 = bf_6_14.text

sc_6_14 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_6_14.text[-1].strip() == '万':
    shangchuan_6_14 = int(float(sc_6_14.text[:-1])*10000)
else:
    shangchuan_6_14 = sc_6_14.text

name_6_14 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
biaotou6_14 = "此次" + name_6_14.text + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou6_14)
print('账号为：',name_6_14.text + " ；上传量：",shangchuan_6_14 , ' ；播放量：' , bofang_6_14 , " ；粉丝量：" ,fensi_6_14)

worksheet.write('A44', name_6_14.text,geshi)
worksheet.write('B44', shangchuan_6_14,geshi)
worksheet.write('C44', bofang_6_14,geshi)
worksheet.write('D44', fensi_6_14,geshi)
worksheet.write('E44', '专题节目部',geshi)

#爱在乡村
id_6_15 = '23442671882004589'
url_6_15 = 'https://m.yangshipin.cn/user?cpid=' + id_6_15
Browser.get(url_6_15)
time.sleep(4)
fs_6_15 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_6_15.text[-1].strip() == '万':
    fensi_6_15 = int(float(fs_6_15.text[:-1])*10000)
else:
    fensi_6_15 = fs_6_15.text

bf_6_15 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_6_15.text[-1].strip() == '万':
    bofang_6_15 = int(float(bf_6_15.text[:-1])*10000)
else:
    bofang_6_15 = bf_6_15.text

sc_6_15 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_6_15.text[-1].strip() == '万':
    shangchuan_6_15 = int(float(sc_6_15.text[:-1])*10000)
else:
    shangchuan_6_15 = sc_6_15.text

name_6_15 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
biaotou6_15 = "此次" + name_6_15.text + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou6_15)
print('账号为：',name_6_15.text + " ；上传量：",shangchuan_6_15 , ' ；播放量：' , bofang_6_15 , " ；粉丝量：" ,fensi_6_15)

worksheet.write('A45', name_6_15.text,geshi)
worksheet.write('B45', shangchuan_6_15,geshi)
worksheet.write('C45', bofang_6_15,geshi)
worksheet.write('D45', fensi_6_15,geshi)
worksheet.write('E45', '合作节目部',geshi)

#大国农数（央视频）
id_6_16 = '30300097029688512'
url_6_16 = 'https://m.yangshipin.cn/user?cpid=' + id_6_16
Browser.get(url_6_16)
time.sleep(4)
fs_6_16 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_6_16.text[-1].strip() == '万':
    fensi_6_16 = int(float(fs_6_16.text[:-1])*10000)
else:
    fensi_6_16 = fs_6_16.text

bf_6_16 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_6_16.text[-1].strip() == '万':
    bofang_6_16 = int(float(bf_6_16.text[:-1])*10000)
else:
    bofang_6_16 = bf_6_16.text

sc_6_16 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_6_16.text[-1].strip() == '万':
    shangchuan_6_16 = int(float(sc_6_16.text[:-1])*10000)
else:
    shangchuan_6_16 = sc_6_16.text

name_6_16_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
name_6_16 = name_6_16_0.text + '(央视频)'
biaotou6_16 = "此次" + name_6_16 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou6_16)
print('账号为：',name_6_16 + " ；上传量：",shangchuan_6_16 , ' ；播放量：' , bofang_6_16 , " ；粉丝量：" ,fensi_6_16)

worksheet.write('A46', name_6_16,geshi)
worksheet.write('B46', shangchuan_6_16,geshi)
worksheet.write('C46', bofang_6_16,geshi)
worksheet.write('D46', fensi_6_16,geshi)
worksheet.write('E46', '广播节目部',geshi)

#乡音博物馆
id_6_17 = '28934401448358814'
url_6_17 = 'https://m.yangshipin.cn/user?cpid=' + id_6_17
Browser.get(url_6_17)
time.sleep(4)
fs_6_17 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_6_17.text[-1].strip() == '万':
    fensi_6_17 = int(float(fs_6_17.text[:-1])*10000)
else:
    fensi_6_17 = fs_6_17.text

bf_6_17 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_6_17.text[-1].strip() == '万':
    bofang_6_17 = int(float(bf_6_17.text[:-1])*10000)
else:
    bofang_6_17 = bf_6_17.text

sc_6_17 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_6_17.text[-1].strip() == '万':
    shangchuan_6_17 = int(float(sc_6_17.text[:-1])*10000)
else:
    shangchuan_6_17 = sc_6_17.text

name_6_17_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
name_6_17 = name_6_17_0.text + '(央视频)'
biaotou6_17 = "此次" + name_6_17 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou6_17)
print('账号为：',name_6_17 + " ；上传量：",shangchuan_6_17 , ' ；播放量：' , bofang_6_17 , " ；粉丝量：" ,fensi_6_17)

worksheet.write('A47', name_6_17,geshi)
worksheet.write('B47', shangchuan_6_17,geshi)
worksheet.write('C47', bofang_6_17,geshi)
worksheet.write('D47', fensi_6_17,geshi)
worksheet.write('E47', '融媒体部',geshi)

#大医生来了（央视频）
id_6_20 = '31516025256067017'
url_6_20 = 'https://m.yangshipin.cn/user?cpid=' + id_6_20
Browser.get(url_6_20)
time.sleep(4)
fs_6_20 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_6_20.text[-1].strip() == '万':
    fensi_6_20 = int(float(fs_6_20.text[:-1])*10000)
else:
    fensi_6_20 = fs_6_20.text

bf_6_20 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_6_20.text[-1].strip() == '万':
    bofang_6_20 = int(float(bf_6_20.text[:-1])*10000)
else:
    bofang_6_20 = bf_6_20.text

sc_6_20 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_6_20.text[-1].strip() == '万':
    shangchuan_6_20 = int(float(sc_6_20.text[:-1])*10000)
else:
    shangchuan_6_20 = sc_6_20.text

name_6_20_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
name_6_20 = name_6_20_0.text + '(央视频)'
biaotou6_20 = "此次" + name_6_20 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou6_20)
print('账号为：',name_6_20 + " ；上传量：",shangchuan_6_20 , ' ；播放量：' , bofang_6_20 , " ；粉丝量：" ,fensi_6_20)

worksheet.write('A48', name_6_20,geshi)
worksheet.write('B48', shangchuan_6_20,geshi)
worksheet.write('C48', bofang_6_20,geshi)
worksheet.write('D48', fensi_6_20,geshi)
worksheet.write('E48', '广播节目部',geshi)

#三农就业行（央视频）
id_6_30 = '28400029167078423'
url_6_30 = 'https://m.yangshipin.cn/user?cpid=' + id_6_30
Browser.get(url_6_30)
time.sleep(4)
fs_6_30 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_6_30.text[-1].strip() == '万':
    fensi_6_30 = int(float(fs_6_30.text[:-1])*10000)
else:
    fensi_6_30 = fs_6_30.text

bf_6_30 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_6_30.text[-1].strip() == '万':
    bofang_6_30 = int(float(bf_6_30.text[:-1])*10000)
else:
    bofang_6_30 = bf_6_30.text

sc_6_30 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_6_30.text[-1].strip() == '万':
    shangchuan_6_30 = int(float(sc_6_30.text[:-1])*10000)
else:
    shangchuan_6_30 = sc_6_30.text

name_6_30_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
name_6_30 = name_6_30_0.text + '(央视频)'
biaotou6_30 = "此次" + name_6_30 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou6_30)
print('账号为：',name_6_30 + " ；上传量：",shangchuan_6_30 , ' ；播放量：' , bofang_6_30 , " ；粉丝量：" ,fensi_6_30)

worksheet.write('A49', name_6_30,geshi)
worksheet.write('B49', shangchuan_6_30,geshi)
worksheet.write('C49', bofang_6_30,geshi)
worksheet.write('D49', fensi_6_30,geshi)
worksheet.write('E49', '项目部',geshi)

#本草华夏（央视频）
id_6_21 = '30300133436341922'
url_6_21 = 'https://m.yangshipin.cn/user?cpid=' + id_6_21
Browser.get(url_6_21)
time.sleep(4)
fs_6_21 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_6_21.text[-1].strip() == '万':
    fensi_6_21 = int(float(fs_6_21.text[:-1])*10000)
else:
    fensi_6_21 = fs_6_21.text

bf_6_21 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_6_21.text[-1].strip() == '万':
    bofang_6_21 = int(float(bf_6_21.text[:-1])*10000)
else:
    bofang_6_21 = bf_6_21.text

sc_6_21 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_6_21.text[-1].strip() == '万':
    shangchuan_6_21 = int(float(sc_6_21.text[:-1])*10000)
else:
    shangchuan_6_21 = sc_6_21.text

name_6_21_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
name_6_21 = name_6_21_0.text + '(央视频)'
biaotou6_21 = "此次" + name_6_21 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou6_21)
print('账号为：',name_6_21 + " ；上传量：",shangchuan_6_21 , ' ；播放量：' , bofang_6_21 , " ；粉丝量：" ,fensi_6_21)

worksheet.write('A50', name_6_21,geshi)
worksheet.write('B50', shangchuan_6_21,geshi)
worksheet.write('C50', bofang_6_21,geshi)
worksheet.write('D50', fensi_6_21,geshi)
worksheet.write('E50', '广播节目部',geshi)

#开箱实验室（央视频）
id_6_22 = '30300060623188800'
url_6_22 = 'https://m.yangshipin.cn/user?cpid=' + id_6_22
Browser.get(url_6_22)
time.sleep(4)
fs_6_22 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_6_22.text[-1].strip() == '万':
    fensi_6_22 = int(float(fs_6_22.text[:-1])*10000)
else:
    fensi_6_22 = fs_6_22.text

bf_6_22 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_6_22.text[-1].strip() == '万':
    bofang_6_22 = int(float(bf_6_22.text[:-1])*10000)
else:
    bofang_6_22 = bf_6_22.text

sc_6_22 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_6_22.text[-1].strip() == '万':
    shangchuan_6_22 = int(float(sc_6_22.text[:-1])*10000)
else:
    shangchuan_6_22 = sc_6_22.text

name_6_22_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
name_6_22 = name_6_22_0.text + '(央视频)'
biaotou6_22 = "此次" + name_6_22 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou6_22)
print('账号为：',name_6_22 + " ；上传量：",shangchuan_6_22 , ' ；播放量：' , bofang_6_22 , " ；粉丝量：" ,fensi_6_22)

worksheet.write('A51', name_6_22,geshi)
worksheet.write('B51', shangchuan_6_22,geshi)
worksheet.write('C51', bofang_6_22,geshi)
worksheet.write('D51', fensi_6_22,geshi)
worksheet.write('E51', '广播节目部',geshi)

#网络丰晚（央视频）
id_6_23 = '24254765472822353'
url_6_23 = 'https://m.yangshipin.cn/user?cpid=' + id_6_23
Browser.get(url_6_23)
time.sleep(4)
fs_6_23 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_6_23.text[-1].strip() == '万':
    fensi_6_23 = int(float(fs_6_23.text[:-1])*10000)
else:
    fensi_6_23 = fs_6_23.text

bf_6_23 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_6_23.text[-1].strip() == '万':
    bofang_6_23 = int(float(bf_6_23.text[:-1])*10000)
else:
    bofang_6_23 = bf_6_23.text

sc_6_23 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_6_23.text[-1].strip() == '万':
    shangchuan_6_23 = int(float(sc_6_23.text[:-1])*10000)
else:
    shangchuan_6_23 = sc_6_23.text

name_6_23_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
name_6_23 = name_6_23_0.text + '(央视频)'
biaotou6_22 = "此次" + name_6_23 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou6_22)
print('账号为：',name_6_23 + " ；上传量：",shangchuan_6_23 , ' ；播放量：' , bofang_6_23 , " ；粉丝量：" ,fensi_6_23)

worksheet.write('A52', name_6_23,geshi)
worksheet.write('B52', shangchuan_6_23,geshi)
worksheet.write('C52', bofang_6_23,geshi)
worksheet.write('D52', fensi_6_23,geshi)
worksheet.write('E52', '项目部',geshi)


#新农研究院（央视频）
id_6_24 = '28863841175393820'
url_6_24 = 'https://m.yangshipin.cn/user?cpid=' + id_6_24
Browser.get(url_6_24)
time.sleep(4)
fs_6_24 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_6_24.text[-1].strip() == '万':
    fensi_6_24 = int(float(fs_6_24.text[:-1])*10000)
else:
    fensi_6_24 = fs_6_24.text

bf_6_24 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_6_24.text[-1].strip() == '万':
    bofang_6_24 = int(float(bf_6_24.text[:-1])*10000)
else:
    bofang_6_24 = bf_6_24.text

sc_6_24 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_6_24.text[-1].strip() == '万':
    shangchuan_6_24 = int(float(sc_6_24.text[:-1])*10000)
else:
    shangchuan_6_24 = sc_6_24.text

name_6_24_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
name_6_24 = name_6_24_0.text + '(央视频)'
biaotou6_22 = "此次" + name_6_24 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou6_22)
print('账号为：',name_6_24 + " ；上传量：",shangchuan_6_24 , ' ；播放量：' , bofang_6_24 , " ；粉丝量：" ,fensi_6_24)

worksheet.write('A53', name_6_24,geshi)
worksheet.write('B53', shangchuan_6_24,geshi)
worksheet.write('C53', bofang_6_24,geshi)
worksheet.write('D53', fensi_6_24,geshi)
worksheet.write('E53', '统筹策划部',geshi)

#中国乡村之声（央视频）
id_6_18 = '31427200802416640'
url_6_18 = 'https://m.yangshipin.cn/user?cpid=' + id_6_18
Browser.get(url_6_18)
time.sleep(4)
fs_6_18 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_6_18.text[-1].strip() == '万':
    fensi_6_18 = int(float(fs_6_18.text[:-1])*10000)
else:
    fensi_6_18 = fs_6_18.text

bf_6_18 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_6_18.text[-1].strip() == '万':
    bofang_6_18 = int(float(bf_6_18.text[:-1])*10000)
else:
    bofang_6_18 = bf_6_18.text

sc_6_18 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_6_18.text[-1].strip() == '万':
    shangchuan_6_18 = int(float(sc_6_18.text[:-1])*10000)
else:
    shangchuan_6_18 = sc_6_18.text

name_6_18_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
name_6_18 = name_6_18_0.text + '(央视频)'
biaotou6_16 = "此次" + name_6_18 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou6_16)
print('账号为：',name_6_18 + " ；上传量：",shangchuan_6_18 , ' ；播放量：' , bofang_6_18 , " ；粉丝量：" ,fensi_6_18)

worksheet.write('A54', name_6_18,geshi)
worksheet.write('B54', shangchuan_6_18,geshi)
worksheet.write('C54', bofang_6_18,geshi)
worksheet.write('D54', fensi_6_18,geshi)
worksheet.write('E54', '广播节目部',geshi)

#茶香中国（央视频）
id_6_19 = '24254827045204566'
url_6_19 = 'https://m.yangshipin.cn/user?cpid=' + id_6_19
Browser.get(url_6_19)
time.sleep(4)
fs_6_19 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_6_19.text[-1].strip() == '万':
    fensi_6_19 = int(float(fs_6_19.text[:-1])*10000)
else:
    fensi_6_19 = fs_6_19.text

bf_6_19 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_6_19.text[-1].strip() == '万':
    bofang_6_19 = int(float(bf_6_19.text[:-1])*10000)
else:
    bofang_6_19 = bf_6_19.text

sc_6_19 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_6_19.text[-1].strip() == '万':
    shangchuan_6_19 = int(float(sc_6_19.text[:-1])*10000)
else:
    shangchuan_6_19 = sc_6_19.text

name_6_19_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
name_6_19 = name_6_19_0.text + '(央视频)'
biaotou6_16 = "此次" + name_6_19 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou6_16)
print('账号为：',name_6_19 + " ；上传量：",shangchuan_6_19 , ' ；播放量：' , bofang_6_19 , " ；粉丝量：" ,fensi_6_19)

worksheet.write('A55', name_6_19,geshi)
worksheet.write('B55', shangchuan_6_19,geshi)
worksheet.write('C55', bofang_6_19,geshi)
worksheet.write('D55', fensi_6_19,geshi)
worksheet.write('E55', '融媒体部',geshi)

#村ba（央视频）
id_6_25 = '38198208188265301'
url_6_25 = 'https://m.yangshipin.cn/user?cpid=' + id_6_25
Browser.get(url_6_25)
time.sleep(4)
fs_6_25 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_6_25.text[-1].strip() == '万':
    fensi_6_25 = int(float(fs_6_25.text[:-1])*10000)
else:
    fensi_6_25 = fs_6_25.text

bf_6_25 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_6_25.text[-1].strip() == '万':
    bofang_6_25 = int(float(bf_6_25.text[:-1])*10000)
else:
    bofang_6_25 = bf_6_25.text

sc_6_25 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_6_25.text[-1].strip() == '万':
    shangchuan_6_25 = int(float(sc_6_25.text[:-1])*10000)
else:
    shangchuan_6_25 = sc_6_25.text

name_6_25_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
name_6_25 = name_6_25_0.text + '(央视频)'
biaotou6_25 = "此次" + name_6_25 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou6_25)
print('账号为：',name_6_25 + " ；上传量：",shangchuan_6_25 , ' ；播放量：' , bofang_6_25 , " ；粉丝量：" ,fensi_6_25)

worksheet.write('A56', name_6_25,geshi)
worksheet.write('B56', shangchuan_6_25,geshi)
worksheet.write('C56', bofang_6_25,geshi)
worksheet.write('D56', fensi_6_25,geshi)
worksheet.write('E56', '融媒体部',geshi)


#超级农人秀（央视频）
id_100_1 = '23598782366960840'
url_100_1 = 'https://m.yangshipin.cn/user?cpid=' + id_100_1
Browser.get(url_100_1)
time.sleep(4)
fs_100_1 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_100_1.text[-1].strip() == '万':
    fensi_100_1 = int(float(fs_100_1.text[:-1]) * 10000)
else:
    fensi_100_1 = fs_100_1.text

bf_100_1 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_100_1.text[-1].strip() == '万':
    bofang_100_1 = int(float(bf_100_1.text[:-1]) * 10000)
else:
    bofang_100_1 = bf_100_1.text

sc_100_1 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_100_1.text[-1].strip() == '万':
    shangchuan_100_1 = int(float(sc_100_1.text[:-1]) * 10000)
else:
    shangchuan_100_1 = sc_100_1.text

name_100_1_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
name_100_1 = name_100_1_0.text + '(央视频)'
biaotou100_1 = "此次" + name_100_1 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ', time.localtime(time.time()))

print(biaotou100_1)
print('账号为：', name_100_1 + " ；上传量：", shangchuan_100_1, ' ；播放量：', bofang_100_1, " ；粉丝量：", fensi_100_1)

worksheet.write('A57', name_100_1, geshi)
worksheet.write('B57', shangchuan_100_1, geshi)
worksheet.write('C57', bofang_100_1, geshi)
worksheet.write('D57', fensi_100_1, geshi)
worksheet.write('E57', '项目部', geshi)

#预制菜争霸赛（央视频）
id_100_2 = '23598803170708751'
url_100_2 = 'https://m.yangshipin.cn/user?cpid=' + id_100_2
Browser.get(url_100_2)
time.sleep(4)
fs_100_2 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_100_2.text[-1].strip() == '万':
    fensi_100_2 = int(float(fs_100_2.text[:-1]) * 10000)
else:
    fensi_100_2 = fs_100_2.text

bf_100_2 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_100_2.text[-1].strip() == '万':
    bofang_100_2 = int(float(bf_100_2.text[:-1]) * 10000)
else:
    bofang_100_2 = bf_100_2.text

sc_100_2 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_100_2.text[-1].strip() == '万':
    shangchuan_100_2 = int(float(sc_100_2.text[:-1]) * 10000)
else:
    shangchuan_100_2 = sc_100_2.text

name_100_2_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
name_100_2 = name_100_2_0.text + '(央视频)'
biaotou100_2 = "此次" + name_100_2 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ', time.localtime(time.time()))

print(biaotou100_2)
print('账号为：', name_100_2 + " ；上传量：", shangchuan_100_2, ' ；播放量：', bofang_100_2, " ；粉丝量：", fensi_100_2)

worksheet.write('A58', name_100_2, geshi)
worksheet.write('B58', shangchuan_100_2, geshi)
worksheet.write('C58', bofang_100_2, geshi)
worksheet.write('D58', fensi_100_2, geshi)
worksheet.write('E58', '项目部', geshi)

#村庄里的中国（央视频）
id_100_3 = '42777257401067599'
url_100_3 = 'https://m.yangshipin.cn/user?cpid=' + id_100_3
Browser.get(url_100_3)
time.sleep(4)
fs_100_3 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_100_3.text[-1].strip() == '万':
    fensi_100_3 = int(float(fs_100_3.text[:-1])*10000)
else:
    fensi_100_3 = fs_100_3.text

bf_100_3 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_100_3.text[-1].strip() == '万':
    bofang_100_3 = int(float(bf_100_3.text[:-1])*10000)
else:
    bofang_100_3 = bf_100_3.text

sc_100_3 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_100_3.text[-1].strip() == '万':
    shangchuan_100_3 = int(float(sc_100_3.text[:-1])*10000)
else:
    shangchuan_100_3 = sc_100_3.text

name_100_3_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
name_100_3 = name_100_3_0.text + '(央视频)'
biaotou100_3 = "此次" + name_100_3 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou100_3)
print('账号为：',name_100_3 + " ；上传量：",shangchuan_100_3 , ' ；播放量：' , bofang_100_3 , " ；粉丝量：" ,fensi_100_3)

worksheet.write('A59', name_100_3,geshi)
worksheet.write('B59', shangchuan_100_3,geshi)
worksheet.write('C59', bofang_100_3,geshi)
worksheet.write('D59', fensi_100_3,geshi)
worksheet.write('E59', '专题节目部',geshi)

#美丽庭院看乡村（央视频）
id_100_4 = '23598734887534166'
url_100_4 = 'https://m.yangshipin.cn/user?cpid=' + id_100_4
Browser.get(url_100_4)
time.sleep(4)
fs_100_4 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_100_4.text[-1].strip() == '万':
    fensi_100_4 = int(float(fs_100_4.text[:-1])*10000)
else:
    fensi_100_4 = fs_100_4.text

bf_100_4 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_100_4.text[-1].strip() == '万':
    bofang_100_4 = int(float(bf_100_4.text[:-1])*10000)
else:
    bofang_100_4 = bf_100_4.text

sc_100_4 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_100_4.text[-1].strip() == '万':
    shangchuan_100_4 = int(float(sc_100_4.text[:-1])*10000)
else:
    shangchuan_100_4 = sc_100_4.text

name_100_4_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
name_100_4 = name_100_4_0.text + '(央视频)'
biaotou100_4 = "此次" + name_100_4 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou100_4)
print('账号为：',name_100_4 + " ；上传量：",shangchuan_100_4 , ' ；播放量：' , bofang_100_4 , " ；粉丝量：" ,fensi_100_4)

worksheet.write('A60', name_100_4,geshi)
worksheet.write('B60', shangchuan_100_4,geshi)
worksheet.write('C60', bofang_100_4,geshi)
worksheet.write('D60', fensi_100_4,geshi)
worksheet.write('E60', '项目部',geshi)


#中国乡土宝藏（央视频）
id_100_5 = '24254672023670309'
url_100_5 = 'https://m.yangshipin.cn/user?cpid=' + id_100_5
Browser.get(url_100_5)
time.sleep(4)
fs_100_5 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_100_5.text[-1].strip() == '万':
    fensi_100_5 = int(float(fs_100_5.text[:-1])*10000)
else:
    fensi_100_5 = fs_100_5.text

bf_100_5 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_100_5.text[-1].strip() == '万':
    bofang_100_5 = int(float(bf_100_5.text[:-1])*10000)
else:
    bofang_100_5 = bf_100_5.text

sc_100_5 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_100_5.text[-1].strip() == '万':
    shangchuan_100_5 = int(float(sc_100_5.text[:-1])*10000)
else:
    shangchuan_100_5 = sc_100_5.text

name_100_5_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
name_100_5 = name_100_5_0.text + '(央视频)'
biaotou100_5 = "此次" + name_100_5 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou100_5)
print('账号为：',name_100_5 + " ；上传量：",shangchuan_100_5 , ' ；播放量：' , bofang_100_5 , " ；粉丝量：" ,fensi_100_5)

worksheet.write('A61', name_100_5,geshi)
worksheet.write('B61', shangchuan_100_5,geshi)
worksheet.write('C61', bofang_100_5,geshi)
worksheet.write('D61', fensi_100_5,geshi)
worksheet.write('E61', '项目部',geshi)


#中国网红村大会（央视频）
id_100_6 = '23598767267466567'
url_100_6 = 'https://m.yangshipin.cn/user?cpid=' + id_100_6
Browser.get(url_100_6)
time.sleep(4)
fs_100_6 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_100_6.text[-1].strip() == '万':
    fensi_100_6 = int(float(fs_100_6.text[:-1])*10000)
else:
    fensi_100_6 = fs_100_6.text

bf_100_6 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_100_6.text[-1].strip() == '万':
    bofang_100_6 = int(float(bf_100_6.text[:-1])*10000)
else:
    bofang_100_6 = bf_100_6.text

sc_100_6 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_100_6.text[-1].strip() == '万':
    shangchuan_100_6 = int(float(sc_100_6.text[:-1])*10000)
else:
    shangchuan_100_6 = sc_100_6.text

name_100_6_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
name_100_6 = name_100_6_0.text + '(央视频)'
biaotou100_6 = "此次" + name_100_6 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou100_6)
print('账号为：',name_100_6 + " ；上传量：",shangchuan_100_6 , ' ；播放量：' , bofang_100_6 , " ；粉丝量：" ,fensi_100_6)

worksheet.write('A62', name_100_6,geshi)
worksheet.write('B62', shangchuan_100_6,geshi)
worksheet.write('C62', bofang_100_6,geshi)
worksheet.write('D62', fensi_100_6,geshi)
worksheet.write('E62', '项目部',geshi)


#丰收地图（央视频）
id_100_7 = '37213172538457633'
url_100_7 = 'https://m.yangshipin.cn/user?cpid=' + id_100_7
Browser.get(url_100_7)
time.sleep(4)
fs_100_7 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_100_7.text[-1].strip() == '万':
    fensi_100_7 = int(float(fs_100_7.text[:-1])*10000)
else:
    fensi_100_7 = fs_100_7.text

bf_100_7 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_100_7.text[-1].strip() == '万':
    bofang_100_7 = int(float(bf_100_7.text[:-1])*10000)
else:
    bofang_100_7 = bf_100_7.text

sc_100_7 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_100_7.text[-1].strip() == '万':
    shangchuan_100_7 = int(float(sc_100_7.text[:-1])*10000)
else:
    shangchuan_100_7 = sc_100_7.text

name_100_7_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
name_100_7 = name_100_7_0.text + '(央视频)'
biaotou100_7 = "此次" + name_100_7 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou100_7)
print('账号为：',name_100_7 + " ；上传量：",shangchuan_100_7 , ' ；播放量：' , bofang_100_7 , " ；粉丝量：" ,fensi_100_7)

worksheet.write('A63', name_100_7,geshi)
worksheet.write('B63', shangchuan_100_7,geshi)
worksheet.write('C63', bofang_100_7,geshi)
worksheet.write('D63', fensi_100_7,geshi)
worksheet.write('E63', '社会交流部',geshi)

#中国三农报道（央视频）
id_100_8 = '46418691232573563'
url_100_8 = 'https://m.yangshipin.cn/user?cpid=' + id_100_8
Browser.get(url_100_8)
time.sleep(4)
fs_100_8 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_100_8.text[-1].strip() == '万':
    fensi_100_8 = int(float(fs_100_8.text[:-1])*10000)
else:
    fensi_100_8 = fs_100_8.text

bf_100_8 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_100_8.text[-1].strip() == '万':
    bofang_100_8 = int(float(bf_100_8.text[:-1])*10000)
else:
    bofang_100_8 = bf_100_8.text

sc_100_8 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_100_8.text[-1].strip() == '万':
    shangchuan_100_8 = int(float(sc_100_8.text[:-1])*10000)
else:
    shangchuan_100_8 = sc_100_8.text

name_100_8_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
name_100_8 = name_100_8_0.text + '(央视频)'
biaotou100_8 = "此次" + name_100_8 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou100_8)
print('账号为：',name_100_8 + " ；上传量：",shangchuan_100_8 , ' ；播放量：' , bofang_100_8 , " ；粉丝量：" ,fensi_100_8)

worksheet.write('A64', name_100_8,geshi)
worksheet.write('B64', shangchuan_100_8,geshi)
worksheet.write('C64', bofang_100_8,geshi)
worksheet.write('D64', fensi_100_8,geshi)
worksheet.write('E64', '新闻部',geshi)


#你好童年
id_100_9 = '26864001222502401'
url_100_9 = 'https://w.yangshipin.cn/user?cpid=' + id_100_9
Browser.get(url_100_9)
time.sleep(4)
fs_100_9 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_100_9.text[-1].strip() == '万':
    fensi_100_9 = int(float(fs_100_9.text[:-1])*10000)
else:
    fensi_100_9 = fs_100_9.text

bf_100_9 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_100_9.text[-1].strip() == '万':
    bofang_100_9 = int(float(bf_100_9.text[:-1])*10000)
else:
    bofang_100_9 = bf_100_9.text

sc_100_9 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_100_9.text[-1].strip() == '万':
    shangchuan_100_9 = int(float(sc_100_9.text[:-1])*10000)
else:
    shangchuan_100_9 = sc_100_9.text

name_100_9_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
name_100_9 = name_100_9_0.text + '(央视频)'
biaotou100_9 = "此次" + name_100_9 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou100_9)
print('账号为：',name_100_9 + " ；上传量：",shangchuan_100_9 , ' ；播放量：' , bofang_100_9 , " ；粉丝量：" ,fensi_100_9)

worksheet.write('A65', name_100_9,geshi)
worksheet.write('B65', shangchuan_100_9,geshi)
worksheet.write('C65', bofang_100_9,geshi)
worksheet.write('D65', fensi_100_9,geshi)


#共话山河
id_100_10 = '31226838229242991'
url_100_10 = 'https://m.yangshipin.cn/user?cpid=' + id_100_10
Browser.get(url_100_10)
time.sleep(4)
fs_100_10 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
if fs_100_10.text[-1].strip() == '万':
    fensi_100_10 = int(float(fs_100_10.text[:-1])*10000)
else:
    fensi_100_10 = fs_100_10.text

bf_100_10 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
if bf_100_10.text[-1].strip() == '万':
    bofang_100_10 = int(float(bf_100_10.text[:-1])*10000)
else:
    bofang_100_10 = bf_100_10.text

sc_100_10 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
if sc_100_10.text[-1].strip() == '万':
    shangchuan_100_10 = int(float(sc_100_10.text[:-1])*10000)
else:
    shangchuan_100_10 = sc_100_10.text

name_100_10_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
name_100_10 = name_100_10_0.text + '(央视频)'
biaotou100_10 = "此次" + name_100_10 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou100_10)
print('账号为：',name_100_10 + " ；上传量：",shangchuan_100_10 , ' ；播放量：' , bofang_100_10 , " ；粉丝量：" ,fensi_100_10)

worksheet.write('A66', name_100_10,geshi)
worksheet.write('B66', shangchuan_100_10,geshi)
worksheet.write('C66', bofang_100_10,geshi)
worksheet.write('D66', fensi_100_10,geshi)

#央视频新增 100_11

#---------------------------------------------------------------------------------------------------------以下是云听专辑

#记忆乡愁之传承
id_4_1 = '15682083075249'
url_4_1 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=541450&columnId=' + id_4_1
Browser.get(url_4_1)
time.sleep(4)

bf_4_1 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_1.text[-1].strip() == '万':
    bofang_4_1 = int(float(bf_4_1.text[:-1])*10000)
else:
    bofang_4_1 = bf_4_1.text

sc_4_1 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_1.text[-1].strip() == '集':
    shangchuan_4_1 = int(float(sc_4_1.text[:-1]))
else:
    shangchuan_4_1 = sc_4_1.text

yuanshimingcheng_4_1 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_1 = yuanshimingcheng_4_1.text
biaotou4_1 = "此次端数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_1)
print('账号为：',name_4_1 + " ；上传量：",shangchuan_4_1 , ' ；播放量：' , bofang_4_1 )

worksheet.set_column('H:H',25)
biaotou4_1 = "此次云听端数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))
worksheet.merge_range('H1:J1',biaotou4_1,biaotougeshi)
worksheet.write('H3', name_4_1,geshi)
worksheet.write('I3', shangchuan_4_1,geshi)
worksheet.write('J3', bofang_4_1,geshi)

worksheet.write('H1', biaotou4_1)
worksheet.write('H2', '账号名称', biaotougeshi)
worksheet.write('I2', '上传量', biaotougeshi)
worksheet.write('J2', '播放量',biaotougeshi)

#攻坚日记（云听）
id_2_1_1 = '15952157837590'
url_2_1_1 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1294228&columnId=' + id_2_1_1
Browser.get(url_2_1_1)
time.sleep(4)

bf_2_1_1 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_2_1_1.text[-1].strip() == '万':
    bofang_2_1_1 = int(float(bf_2_1_1.text[:-1])*10000)
else:
    bofang_2_1_1 = bf_2_1_1.text

sc_2_1_1 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_2_1_1.text[-1].strip() == '集':
    shangchuan_2_1_1 = int(float(sc_2_1_1.text[:-1]))
else:
    shangchuan_2_1_1 = sc_2_1_1.text

yuanshimingcheng_2_1_1 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_2_1_1 = yuanshimingcheng_2_1_1.text + '(云听)'
biaotou2_1_1 = "此次" + name_2_1_1 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou2_1_1)
print('账号为：',name_2_1_1 + " ；上传量：",shangchuan_2_1_1 , ' ；播放量：' , bofang_2_1_1 )

worksheet.write('H4', name_2_1_1,geshi)
worksheet.write('I4', shangchuan_2_1_1,geshi)
worksheet.write('J4', bofang_2_1_1,geshi)

#大地讲堂（云听）
id_3_8 = '15955553258740'
url_3_8 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1286791&columnId=' + id_3_8
Browser.get(url_3_8)
time.sleep(7)

bf_3_8 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_3_8.text[-1].strip() == '万':
    bofang_3_8 = int(float(bf_3_8.text[:-1])*10000)
else:
    bofang_3_8 = bf_3_8.text

sc_3_8 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_3_8.text[-1].strip() == '集':
    shangchuan_3_8 = int(float(sc_3_8.text[:-1]))
else:
    shangchuan_3_8 = sc_3_8.text

yuanshimingcheng_3_8 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_3_8 = yuanshimingcheng_3_8.text + '(云听)'
biaotou3_8 = "此次" + name_3_8 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou3_8)
print('账号为：',name_3_8 + " ；上传量：",shangchuan_3_8 , ' ；播放量：' , bofang_3_8 )

worksheet.write('H5', name_3_8,geshi)
worksheet.write('I5', shangchuan_3_8,geshi)
worksheet.write('J5', bofang_3_8,geshi)

#记忆乡愁名家诵读
id_4_4 = '15682083075250'
url_4_4 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=541458&columnId=' + id_4_4
Browser.get(url_4_4)
time.sleep(4)

bf_4_4 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_4.text[-1].strip() == '万':
    bofang_4_4 = int(float(bf_4_4.text[:-1])*10000)
else:
    bofang_4_4 = bf_4_4.text

sc_4_4 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_4.text[-1].strip() == '集':
    shangchuan_4_4 = int(float(sc_4_4.text[:-1]))
else:
    shangchuan_4_4 = sc_4_4.text

yuanshimingcheng_4_4 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_4 = yuanshimingcheng_4_4.text
biaotou4_4 = "此次" + name_4_4 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_4)
print('账号为：',name_4_4 + " ；上传量：",shangchuan_4_4 , ' ；播放量：' , bofang_4_4 )

worksheet.write('H6', name_4_4,geshi)
worksheet.write('I6', shangchuan_4_4,geshi)
worksheet.write('J6', bofang_4_4,geshi)

# 我给爸妈唱支歌
id_200_49 = '15682080945904'
url_200_49 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=520838&columnId=' + id_200_49
Browser.get(url_200_49)
time.sleep(4)

bf_200_49 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_49.text[-1].strip() == '万':
    bofang_200_49 = int(float(bf_200_49.text[:-1])*10000)
else:
    bofang_200_49 = bf_200_49.text

sc_200_49 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_49.text[-1].strip() == '集':
    shangchuan_200_49 = int(float(sc_200_49.text[:-1]))
else:
    shangchuan_200_49 = sc_200_49.text

yuanshimingcheng_200_49 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_49 = yuanshimingcheng_200_49.text
biaotou200_49 = "此次" + name_200_49 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_49)
print('账号为：',name_200_49 + " ；上传量：",shangchuan_200_49 , ' ；播放量：' , bofang_200_49 )

worksheet.write('H7', name_200_49,geshi)
worksheet.write('I7', shangchuan_200_49,geshi)
worksheet.write('J7', bofang_200_49,geshi)

#【大地微音】系列微广播剧
id_4_6 = '15682083075324'
url_4_6 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=551929&columnId=' + id_4_6
Browser.get(url_4_6)
time.sleep(4)

bf_4_6 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_6.text[-1].strip() == '万':
    bofang_4_6 = int(float(bf_4_6.text[:-1])*10000)
else:
    bofang_4_6 = bf_4_6.text

sc_4_6 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_6.text[-1].strip() == '集':
    shangchuan_4_6 = int(float(sc_4_6.text[:-1]))
else:
    shangchuan_4_6 = sc_4_6.text

yuanshimingcheng_4_6 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_6 = yuanshimingcheng_4_6.text
biaotou4_6 = "此次" + name_4_6 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_6)
print('账号为：',name_4_6 + " ；上传量：",shangchuan_4_6 , ' ；播放量：' , bofang_4_6 )

worksheet.write('H8', name_4_6,geshi)
worksheet.write('I8', shangchuan_4_6,geshi)
worksheet.write('J8', bofang_4_6,geshi)

#举案说法
id_4_7 = '15682083075333'
url_4_7 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=552129&columnId=' + id_4_7
Browser.get(url_4_7)
time.sleep(4)

bf_4_7 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_7.text[-1].strip() == '万':
    bofang_4_7 = int(float(bf_4_7.text[:-1])*10000)
else:
    bofang_4_7 = bf_4_7.text

sc_4_7 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_7.text[-1].strip() == '集':
    shangchuan_4_7 = int(float(sc_4_7.text[:-1]))
else:
    shangchuan_4_7 = sc_4_7.text

yuanshimingcheng_4_7 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_7 = yuanshimingcheng_4_7.text
biaotou4_7 = "此次" + name_4_7 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_7)
print('账号为：',name_4_7 + " ；上传量：",shangchuan_4_7 , ' ；播放量：' , bofang_4_7 )

worksheet.write('H9', name_4_7,geshi)
worksheet.write('I9', shangchuan_4_7,geshi)
worksheet.write('J9', bofang_4_7,geshi)

#农村土地法律知识30问
id_4_8 = '15682083075349'
url_4_8 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=552743&columnId=' + id_4_8
Browser.get(url_4_8)
time.sleep(4)

bf_4_8 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_8.text[-1].strip() == '万':
    bofang_4_8 = int(float(bf_4_8.text[:-1])*10000)
else:
    bofang_4_8 = bf_4_8.text

sc_4_8 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_8.text[-1].strip() == '集':
    shangchuan_4_8 = int(float(sc_4_8.text[:-1]))
else:
    shangchuan_4_8 = sc_4_8.text

yuanshimingcheng_4_8 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_8 = yuanshimingcheng_4_8.text
biaotou4_8 = "此次" + name_4_8 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_8)
print('账号为：',name_4_8 + " ；上传量：",shangchuan_4_8 , ' ；播放量：' , bofang_4_8 )

worksheet.write('H10', name_4_8,geshi)
worksheet.write('I10', shangchuan_4_8,geshi)
worksheet.write('J10', bofang_4_8,geshi)

#凡事有理
id_4_9 = '16070884252960'
url_4_9 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1373873&columnId=' + id_4_9
Browser.get(url_4_9)
time.sleep(4)

bf_4_9 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_9.text[-1].strip() == '万':
    bofang_4_9 = int(float(bf_4_9.text[:-1])*10000)
else:
    bofang_4_9 = bf_4_9.text

sc_4_9 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_9.text[-1].strip() == '集':
    shangchuan_4_9 = int(float(sc_4_9.text[:-1]))
else:
    shangchuan_4_9 = sc_4_9.text

yuanshimingcheng_4_9 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_9 = yuanshimingcheng_4_9.text
biaotou4_9 = "此次" + name_4_9 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_9)
print('账号为：',name_4_9 + " ；上传量：",shangchuan_4_9 , ' ；播放量：' , bofang_4_9 )

worksheet.write('H11', name_4_9,geshi)
worksheet.write('I11', shangchuan_4_9,geshi)
worksheet.write('J11', bofang_4_9,geshi)

#家书——温暖的力量
id_4_10 = '15682083075385'
url_4_10 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=554556&columnId=' + id_4_10
Browser.get(url_4_10)
time.sleep(4)

bf_4_10 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_10.text[-1].strip() == '万':
    bofang_4_10 = int(float(bf_4_10.text[:-1])*10000)
else:
    bofang_4_10 = bf_4_10.text

sc_4_10 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_10.text[-1].strip() == '集':
    shangchuan_4_10 = int(float(sc_4_10.text[:-1]))
else:
    shangchuan_4_10 = sc_4_10.text

yuanshimingcheng_4_10 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_10 = yuanshimingcheng_4_10.text
biaotou4_10 = "此次" + name_4_10 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_10)
print('账号为：',name_4_10 + " ；上传量：",shangchuan_4_10 , ' ；播放量：' , bofang_4_10 )

worksheet.write('H12', name_4_10,geshi)
worksheet.write('I12', shangchuan_4_10,geshi)
worksheet.write('J12', bofang_4_10,geshi)

#家书——远方的思念
id_4_11 = '15682083075386'
url_4_11 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=554577&columnId=' + id_4_11
Browser.get(url_4_11)
time.sleep(4)

bf_4_11 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_11.text[-1].strip() == '万':
    bofang_4_11 = int(float(bf_4_11.text[:-1])*10000)
else:
    bofang_4_11 = bf_4_11.text

sc_4_11 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_11.text[-1].strip() == '集':
    shangchuan_4_11 = int(float(sc_4_11.text[:-1]))
else:
    shangchuan_4_11 = sc_4_11.text

yuanshimingcheng_4_11 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_11 = yuanshimingcheng_4_11.text
biaotou4_11 = "此次" + name_4_11 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_11)
print('账号为：',name_4_11 + " ；上传量：",shangchuan_4_11 , ' ；播放量：' , bofang_4_11 )

worksheet.write('H13', name_4_11,geshi)
worksheet.write('I13', shangchuan_4_11,geshi)
worksheet.write('J13', bofang_4_11,geshi)

# 铲屎官修炼宝典
id_200_48 = '16003280846390'
url_200_48 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1317764&columnId=' + id_200_48
Browser.get(url_200_48)
time.sleep(4)

bf_200_48 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_48.text[-1].strip() == '万':
    bofang_200_48 = int(float(bf_200_48.text[:-1])*10000)
else:
    bofang_200_48 = bf_200_48.text

sc_200_48 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_48.text[-1].strip() == '集':
    shangchuan_200_48 = int(float(sc_200_48.text[:-1]))
else:
    shangchuan_200_48 = sc_200_48.text

yuanshimingcheng_200_48 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_48 = yuanshimingcheng_200_48.text
biaotou200_48 = "此次" + name_200_48 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_48)
print('账号为：',name_200_48 + " ；上传量：",shangchuan_200_48 , ' ；播放量：' , bofang_200_48 )

worksheet.write('H14', name_200_48,geshi)
worksheet.write('I14', shangchuan_200_48,geshi)
worksheet.write('J14', bofang_200_48,geshi)

#爱在乡村
id_4_13 = '15682083075426'
url_4_13 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=555623&columnId=' + id_4_13
Browser.get(url_4_13)
time.sleep(4)

bf_4_13 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_13.text[-1].strip() == '万':
    bofang_4_13 = int(float(bf_4_13.text[:-1])*10000)
else:
    bofang_4_13 = bf_4_13.text

sc_4_13 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_13.text[-1].strip() == '集':
    shangchuan_4_13 = int(float(sc_4_13.text[:-1]))
else:
    shangchuan_4_13 = sc_4_13.text

yuanshimingcheng_4_13 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_13 = yuanshimingcheng_4_13.text
biaotou4_13 = "此次" + name_4_13 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_13)
print('账号为：',name_4_13 + " ；上传量：",shangchuan_4_13 , ' ；播放量：' , bofang_4_13 )

worksheet.write('H15', name_4_13,geshi)
worksheet.write('I15', shangchuan_4_13,geshi)
worksheet.write('J15', bofang_4_13,geshi)

# 菜市场经济学
id_200_47 = '16509538634580'
url_200_47 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1742127&columnId=' + id_200_47
Browser.get(url_200_47)
time.sleep(4)

bf_200_47 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_47.text[-1].strip() == '万':
    bofang_200_47 = int(float(bf_200_47.text[:-1])*10000)
else:
    bofang_200_47 = bf_200_47.text

sc_200_47 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_47.text[-1].strip() == '集':
    shangchuan_200_47 = int(float(sc_200_47.text[:-1]))
else:
    shangchuan_200_47 = sc_200_47.text

yuanshimingcheng_200_47 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_47 = yuanshimingcheng_200_47.text
biaotou200_47 = "此次" + name_200_47 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_47)
print('账号为：',name_200_47 + " ；上传量：",shangchuan_200_47 , ' ；播放量：' , bofang_200_47 )

worksheet.write('H16', name_200_47,geshi)
worksheet.write('I16', shangchuan_200_47,geshi)
worksheet.write('J16', bofang_200_47,geshi)

#乡村夜话
id_4_15 = '15682083075638'
url_4_15 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=575526&columnId=' + id_4_15
Browser.get(url_4_15)
time.sleep(4)

bf_4_15 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_15.text[-1].strip() == '万':
    bofang_4_15 = int(float(bf_4_15.text[:-1])*10000)
else:
    bofang_4_15 = bf_4_15.text

sc_4_15 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_15.text[-1].strip() == '集':
    shangchuan_4_15 = int(float(sc_4_15.text[:-1]))
else:
    shangchuan_4_15 = sc_4_15.text

yuanshimingcheng_4_15 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_15 = yuanshimingcheng_4_15.text
biaotou4_15 = "此次" + name_4_15 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_15)
print('账号为：',name_4_15 + " ；上传量：",shangchuan_4_15 , ' ；播放量：' , bofang_4_15 )

worksheet.write('H17', name_4_15,geshi)
worksheet.write('I17', shangchuan_4_15,geshi)
worksheet.write('J17', bofang_4_15,geshi)

#乡村保健站
id_4_16 = '15682083075657'
url_4_16 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1118341&columnId=' + id_4_16
Browser.get(url_4_16)
time.sleep(4)

bf_4_16 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_16.text[-1].strip() == '万':
    bofang_4_16 = int(float(bf_4_16.text[:-1])*10000)
else:
    bofang_4_16 = bf_4_16.text

sc_4_16 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_16.text[-1].strip() == '集':
    shangchuan_4_16 = int(float(sc_4_16.text[:-1]))
else:
    shangchuan_4_16 = sc_4_16.text

yuanshimingcheng_4_16 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_16 = yuanshimingcheng_4_16.text
biaotou4_16 = "此次" + name_4_16 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_16)
print('账号为：',name_4_16 + " ；上传量：",shangchuan_4_16 , ' ；播放量：' , bofang_4_16 )

worksheet.write('H18', name_4_16,geshi)
worksheet.write('I18', shangchuan_4_16,geshi)
worksheet.write('J18', bofang_4_16,geshi)

#养生妙招
id_4_17 = '15682083075660'
url_4_17 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1285275&columnId=' + id_4_17
Browser.get(url_4_17)
time.sleep(4)

bf_4_17 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_17.text[-1].strip() == '万':
    bofang_4_17 = int(float(bf_4_17.text[:-1])*10000)
else:
    bofang_4_17 = bf_4_17.text

sc_4_17 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_17.text[-1].strip() == '集':
    shangchuan_4_17 = int(float(sc_4_17.text[:-1]))
else:
    shangchuan_4_17 = sc_4_17.text

yuanshimingcheng_4_17 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_17 = yuanshimingcheng_4_17.text
biaotou4_17 = "此次" + name_4_17 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_17)
print('账号为：',name_4_17 + " ；上传量：",shangchuan_4_17 , ' ；播放量：' , bofang_4_17 )

worksheet.write('H19', name_4_17,geshi)
worksheet.write('I19', shangchuan_4_17,geshi)
worksheet.write('J19', bofang_4_17,geshi)

#大咖讲故事
id_4_18 = '15682083075698'
url_4_18 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=587578&columnId=' + id_4_18
Browser.get(url_4_18)
time.sleep(4)

bf_4_18 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_18.text[-1].strip() == '万':
    bofang_4_18 = int(float(bf_4_18.text[:-1])*10000)
else:
    bofang_4_18 = bf_4_18.text

sc_4_18 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_18.text[-1].strip() == '集':
    shangchuan_4_18 = int(float(sc_4_18.text[:-1]))
else:
    shangchuan_4_18 = sc_4_18.text

yuanshimingcheng_4_18 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_18 = yuanshimingcheng_4_18.text
biaotou4_18 = "此次" + name_4_18 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_18)
print('账号为：',name_4_18 + " ；上传量：",shangchuan_4_18 , ' ；播放量：' , bofang_4_18 )

worksheet.write('H20', name_4_18,geshi)
worksheet.write('I20', shangchuan_4_18,geshi)
worksheet.write('J20', bofang_4_18,geshi)

#名医坐堂
id_4_19 = '15716568375325'
url_4_19 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1282453&columnId=' + id_4_19
Browser.get(url_4_19)
time.sleep(4)

bf_4_19 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_19.text[-1].strip() == '万':
    bofang_4_19 = int(float(bf_4_19.text[:-1])*10000)
else:
    bofang_4_19 = bf_4_19.text

sc_4_19 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_19.text[-1].strip() == '集':
    shangchuan_4_19 = int(float(sc_4_19.text[:-1]))
else:
    shangchuan_4_19 = sc_4_19.text

yuanshimingcheng_4_19 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_19 = yuanshimingcheng_4_19.text
biaotou4_19 = "此次" + name_4_19 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_19)
print('账号为：',name_4_19 + " ；上传量：",shangchuan_4_19 , ' ；播放量：' , bofang_4_19 )

worksheet.write('H21', name_4_19,geshi)
worksheet.write('I21', shangchuan_4_19,geshi)
worksheet.write('J21', bofang_4_19,geshi)

#案例
id_4_20 = '15935815065790'
url_4_20 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1268026&columnId=' + id_4_20
Browser.get(url_4_20)
time.sleep(4)

bf_4_20 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_20.text[-1].strip() == '万':
    bofang_4_20 = int(float(bf_4_20.text[:-1])*10000)
else:
    bofang_4_20 = bf_4_20.text

sc_4_20 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_20.text[-1].strip() == '集':
    shangchuan_4_20 = int(float(sc_4_20.text[:-1]))
else:
    shangchuan_4_20 = sc_4_20.text

yuanshimingcheng_4_20 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_20 = yuanshimingcheng_4_20.text
biaotou4_20 = "此次" + name_4_20 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_20)
print('账号为：',name_4_20 + " ；上传量：",shangchuan_4_20 , ' ；播放量：' , bofang_4_20 )

worksheet.write('H22', name_4_20,geshi)
worksheet.write('I22', shangchuan_4_20,geshi)
worksheet.write('J22', bofang_4_20,geshi)

# 快乐暑假，健康生活
id_4_21 = '15952167682050'
url_4_21 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1289690&columnId=' + id_4_21
Browser.get(url_4_21)
time.sleep(4)

bf_4_21 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_21.text[-1].strip() == '万':
    bofang_4_21 = int(float(bf_4_21.text[:-1])*10000)
else:
    bofang_4_21 = bf_4_21.text

sc_4_21 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_21.text[-1].strip() == '集':
    shangchuan_4_21 = int(float(sc_4_21.text[:-1]))
else:
    shangchuan_4_21 = sc_4_21.text

yuanshimingcheng_4_21 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_21 = yuanshimingcheng_4_21.text
biaotou4_21 = "此次" + name_4_21 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_21)
print('账号为：',name_4_21 + " ；上传量：",shangchuan_4_21 , ' ；播放量：' , bofang_4_21 )

worksheet.write('H23', name_4_21,geshi)
worksheet.write('I23', shangchuan_4_21,geshi)
worksheet.write('J23', bofang_4_21,geshi)

# 乡关何处
id_200_46 = '16039727651670'
url_200_46 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1359442&columnId=' + id_200_46
Browser.get(url_200_46)
time.sleep(4)

bf_200_46 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_46.text[-1].strip() == '万':
    bofang_200_46 = int(float(bf_200_46.text[:-1])*10000)
else:
    bofang_200_46 = bf_200_46.text

sc_200_46 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_46.text[-1].strip() == '集':
    shangchuan_200_46 = int(float(sc_200_46.text[:-1]))
else:
    shangchuan_200_46 = sc_200_46.text

yuanshimingcheng_200_46 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_46 = yuanshimingcheng_200_46.text
biaotou200_46 = "此次" + name_200_46 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_46)
print('账号为：',name_200_46 + " ；上传量：",shangchuan_200_46 , ' ；播放量：' , bofang_200_46 )

worksheet.write('H24', name_200_46,geshi)
worksheet.write('I24', shangchuan_200_46,geshi)
worksheet.write('J24', bofang_200_46,geshi)

# 中国粮食的信心
id_4_23 = '15901307361490'
url_4_23 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1289769&columnId=' + id_4_23
Browser.get(url_4_23)
time.sleep(4)

bf_4_23 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_23.text[-1].strip() == '万':
    bofang_4_23 = int(float(bf_4_23.text[:-1])*10000)
else:
    bofang_4_23 = bf_4_23.text

sc_4_23 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_23.text[-1].strip() == '集':
    shangchuan_4_23 = int(float(sc_4_23.text[:-1]))
else:
    shangchuan_4_23 = sc_4_23.text

yuanshimingcheng_4_23 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_23 = yuanshimingcheng_4_23.text
biaotou4_23 = "此次" + name_4_23 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_23)
print('账号为：',name_4_23 + " ；上传量：",shangchuan_4_23 , ' ；播放量：' , bofang_4_23 )

worksheet.write('H25', name_4_23,geshi)
worksheet.write('I25', shangchuan_4_23,geshi)
worksheet.write('J25', bofang_4_23,geshi)

#三农厨房
id_4_24 = '15779309122640'
url_4_24 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1139627&columnId=' + id_4_24
Browser.get(url_4_24)
time.sleep(4)

bf_4_24 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_24.text[-1].strip() == '万':
    bofang_4_24 = int(float(bf_4_24.text[:-1])*10000)
else:
    bofang_4_24 = bf_4_24.text

sc_4_24 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_24.text[-1].strip() == '集':
    shangchuan_4_24 = int(float(sc_4_24.text[:-1]))
else:
    shangchuan_4_24 = sc_4_24.text

yuanshimingcheng_4_24 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_24 = yuanshimingcheng_4_24.text
biaotou4_24 = "此次" + name_4_24 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_24)
print('账号为：',name_4_24 + " ；上传量：",shangchuan_4_24 , ' ；播放量：' , bofang_4_24 )

worksheet.write('H26', name_4_24,geshi)
worksheet.write('I26', shangchuan_4_24,geshi)
worksheet.write('J26', bofang_4_24,geshi)

#专家们的百变厨坊
id_4_25 = '15919489217940'
url_4_25 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1295106&columnId=' + id_4_25
Browser.get(url_4_25)
time.sleep(4)

bf_4_25 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_25.text[-1].strip() == '万':
    bofang_4_25 = int(float(bf_4_25.text[:-1])*10000)
else:
    bofang_4_25 = bf_4_25.text

sc_4_25 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_25.text[-1].strip() == '集':
    shangchuan_4_25 = int(float(sc_4_25.text[:-1]))
else:
    shangchuan_4_25 = sc_4_25.text

yuanshimingcheng_4_25 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_25 = yuanshimingcheng_4_25.text
biaotou4_25 = "此次" + name_4_25 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_25)
print('账号为：',name_4_25 + " ；上传量：",shangchuan_4_25 , ' ；播放量：' , bofang_4_25 )

worksheet.write('H27', name_4_25,geshi)
worksheet.write('I27', shangchuan_4_25,geshi)
worksheet.write('J27', bofang_4_25,geshi)

#减肥其实很简单
id_4_26 = '15960147270330'
url_4_26 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1295196&columnId=' + id_4_26
Browser.get(url_4_26)
time.sleep(4)

bf_4_26 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_26.text[-1].strip() == '万':
    bofang_4_26 = int(float(bf_4_26.text[:-1])*10000)
else:
    bofang_4_26 = bf_4_26.text

sc_4_26 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_26.text[-1].strip() == '集':
    shangchuan_4_26 = int(float(sc_4_26.text[:-1]))
else:
    shangchuan_4_26 = sc_4_26.text

yuanshimingcheng_4_26 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_26 = yuanshimingcheng_4_26.text
biaotou4_26 = "此次" + name_4_26 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_26)
print('账号为：',name_4_26 + " ；上传量：",shangchuan_4_26 , ' ；播放量：' , bofang_4_26 )

worksheet.write('H28', name_4_26,geshi)
worksheet.write('I28', shangchuan_4_26,geshi)
worksheet.write('J28', bofang_4_26,geshi)

#教你看气象
id_4_27 = '15933250005440'
url_4_27 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1266305&columnId=' + id_4_27
Browser.get(url_4_27)
time.sleep(4)

bf_4_27 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_27.text[-1].strip() == '万':
    bofang_4_27 = int(float(bf_4_27.text[:-1])*10000)
else:
    bofang_4_27 = bf_4_27.text

sc_4_27 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_27.text[-1].strip() == '集':
    shangchuan_4_27 = int(float(sc_4_27.text[:-1]))
else:
    shangchuan_4_27 = sc_4_27.text

yuanshimingcheng_4_27 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_27 = yuanshimingcheng_4_27.text
biaotou4_27 = "此次" + name_4_27 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_27)
print('账号为：',name_4_27 + " ；上传量：",shangchuan_4_27 , ' ；播放量：' , bofang_4_27 )

worksheet.write('H29', name_4_27,geshi)
worksheet.write('I29', shangchuan_4_27,geshi)
worksheet.write('J29', bofang_4_27,geshi)

#三人话三农——对标全面小康
id_4_28 = '15902124283750'
url_4_28 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1211886&columnId=' + id_4_28
Browser.get(url_4_28)
time.sleep(4)

bf_4_28 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_28.text[-1].strip() == '万':
    bofang_4_28 = int(float(bf_4_28.text[:-1])*10000)
else:
    bofang_4_28 = bf_4_28.text

sc_4_28 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_28.text[-1].strip() == '集':
    shangchuan_4_28 = int(float(sc_4_28.text[:-1]))
else:
    shangchuan_4_28 = sc_4_28.text

yuanshimingcheng_4_28 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_28 = yuanshimingcheng_4_28.text
biaotou4_28 = "此次" + name_4_28 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_28)
print('账号为：',name_4_28 + " ；上传量：",shangchuan_4_28 , ' ；播放量：' , bofang_4_28 )

worksheet.write('H30', name_4_28,geshi)
worksheet.write('I30', shangchuan_4_28,geshi)
worksheet.write('J30', bofang_4_28,geshi)

#民法典系列解读
id_4_29 = '15911671773740'
url_4_29 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1294913&columnId=' + id_4_29
Browser.get(url_4_29)
time.sleep(4)

bf_4_29 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_29.text[-1].strip() == '万':
    bofang_4_29 = int(float(bf_4_29.text[:-1])*10000)
else:
    bofang_4_29 = bf_4_29.text

sc_4_29 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_29.text[-1].strip() == '集':
    shangchuan_4_29 = int(float(sc_4_29.text[:-1]))
else:
    shangchuan_4_29 = sc_4_29.text

yuanshimingcheng_4_29 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_29 = yuanshimingcheng_4_29.text
biaotou4_29 = "此次" + name_4_29 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_29)
print('账号为：',name_4_29 + " ；上传量：",shangchuan_4_29 , ' ；播放量：' , bofang_4_29 )

worksheet.write('H31', name_4_29,geshi)
worksheet.write('I31', shangchuan_4_29,geshi)
worksheet.write('J31', bofang_4_29,geshi)

#听海
id_4_30 = '15961688581430'
url_4_30 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1294826&columnId=' + id_4_30
Browser.get(url_4_30)
time.sleep(4)

bf_4_30 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_30.text[-1].strip() == '万':
    bofang_4_30 = int(float(bf_4_30.text[:-1])*10000)
else:
    bofang_4_30 = bf_4_30.text

sc_4_30 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_30.text[-1].strip() == '集':
    shangchuan_4_30 = int(float(sc_4_30.text[:-1]))
else:
    shangchuan_4_30 = sc_4_30.text

yuanshimingcheng_4_30 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_30 = yuanshimingcheng_4_30.text
biaotou4_30 = "此次" + name_4_30 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_30)
print('账号为：',name_4_30 + " ；上传量：",shangchuan_4_30 , ' ；播放量：' , bofang_4_30 )

worksheet.write('H32', name_4_30,geshi)
worksheet.write('I32', shangchuan_4_30,geshi)
worksheet.write('J32', bofang_4_30,geshi)

#乡村振兴资讯
id_4_31 = '15902152391350'
url_4_31 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1211895&columnId=' + id_4_31
Browser.get(url_4_31)
time.sleep(4)

bf_4_31 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_31.text[-1].strip() == '万':
    bofang_4_31 = int(float(bf_4_31.text[:-1])*10000)
else:
    bofang_4_31 = bf_4_31.text

sc_4_31 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_31.text[-1].strip() == '集':
    shangchuan_4_31 = int(float(sc_4_31.text[:-1]))
else:
    shangchuan_4_31 = sc_4_31.text

yuanshimingcheng_4_31 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_31 = yuanshimingcheng_4_31.text
biaotou4_31 = "此次" + name_4_31 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_31)
print('账号为：',name_4_31 + " ；上传量：",shangchuan_4_31 , ' ；播放量：' , bofang_4_31 )

worksheet.write('H33', name_4_31,geshi)
worksheet.write('I33', shangchuan_4_31,geshi)
worksheet.write('J33', bofang_4_31,geshi)

#夏日小花园
id_4_32 = '15955730592130'
url_4_32 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1289717&columnId=' + id_4_32
Browser.get(url_4_32)
time.sleep(4)

bf_4_32 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_32.text[-1].strip() == '万':
    bofang_4_32 = int(float(bf_4_32.text[:-1])*10000)
else:
    bofang_4_32 = bf_4_32.text

sc_4_32 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_32.text[-1].strip() == '集':
    shangchuan_4_32 = int(float(sc_4_32.text[:-1]))
else:
    shangchuan_4_32 = sc_4_32.text

yuanshimingcheng_4_32 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_32 = yuanshimingcheng_4_32.text
biaotou4_32 = "此次" + name_4_32 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_32)
print('账号为：',name_4_32 + " ；上传量：",shangchuan_4_32 , ' ；播放量：' , bofang_4_32 )

worksheet.write('H34', name_4_32,geshi)
worksheet.write('I34', shangchuan_4_32,geshi)
worksheet.write('J34', bofang_4_32,geshi)

#健康到家
id_4_33 = '15707280300500'
url_4_33 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1284650&columnId=' + id_4_33
Browser.get(url_4_33)
time.sleep(4)

bf_4_33 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_33.text[-1].strip() == '万':
    bofang_4_33 = int(float(bf_4_33.text[:-1])*10000)
else:
    bofang_4_33 = bf_4_33.text

sc_4_33 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_33.text[-1].strip() == '集':
    shangchuan_4_33 = int(float(sc_4_33.text[:-1]))
else:
    shangchuan_4_33 = sc_4_33.text

yuanshimingcheng_4_33 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_33 = yuanshimingcheng_4_33.text
biaotou4_33 = "此次" + name_4_33 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_33)
print('账号为：',name_4_33 + " ；上传量：",shangchuan_4_33 , ' ；播放量：' , bofang_4_33 )

worksheet.write('H35', name_4_33,geshi)
worksheet.write('I35', shangchuan_4_33,geshi)
worksheet.write('J35', bofang_4_33,geshi)

#法眼看剧
id_4_34 = '15965388238580'
url_4_34 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1294123&columnId=' + id_4_34
Browser.get(url_4_34)
time.sleep(4)

bf_4_34 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_34.text[-1].strip() == '万':
    bofang_4_34 = int(float(bf_4_34.text[:-1])*10000)
else:
    bofang_4_34 = bf_4_34.text

sc_4_34 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_34.text[-1].strip() == '集':
    shangchuan_4_34 = int(float(sc_4_34.text[:-1]))
else:
    shangchuan_4_34 = sc_4_34.text

yuanshimingcheng_4_34 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_34 = yuanshimingcheng_4_34.text
biaotou4_34 = "此次" + name_4_34 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_34)
print('账号为：',name_4_34 + " ；上传量：",shangchuan_4_34 , ' ；播放量：' , bofang_4_34 )

worksheet.write('H36', name_4_34,geshi)
worksheet.write('I36', shangchuan_4_34,geshi)
worksheet.write('J36', bofang_4_34,geshi)

#第一书记的劳动节
id_4_35 = '15883069260810'
url_4_35 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1205220&columnId=' + id_4_35
Browser.get(url_4_35)
time.sleep(4)

bf_4_35 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_35.text[-1].strip() == '万':
    bofang_4_35 = int(float(bf_4_35.text[:-1])*10000)
else:
    bofang_4_35 = bf_4_35.text

sc_4_35 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_35.text[-1].strip() == '集':
    shangchuan_4_35 = int(float(sc_4_35.text[:-1]))
else:
    shangchuan_4_35 = sc_4_35.text

yuanshimingcheng_4_35 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_35 = yuanshimingcheng_4_35.text
biaotou4_35 = "此次" + name_4_35 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_35)
print('账号为：',name_4_35 + " ；上传量：",shangchuan_4_35 , ' ；播放量：' , bofang_4_35 )

worksheet.write('H37', name_4_35,geshi)
worksheet.write('I37', shangchuan_4_35,geshi)
worksheet.write('J37', bofang_4_35,geshi)

#寻医问诊
id_4_36 = '15888190579250'
url_4_36 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1263850&columnId=' + id_4_36
Browser.get(url_4_36)
time.sleep(4)

bf_4_36 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_36.text[-1].strip() == '万':
    bofang_4_36 = int(float(bf_4_36.text[:-1])*10000)
else:
    bofang_4_36 = bf_4_36.text

sc_4_36 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_36.text[-1].strip() == '集':
    shangchuan_4_36 = int(float(sc_4_36.text[:-1]))
else:
    shangchuan_4_36 = sc_4_36.text

yuanshimingcheng_4_36 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_36 = yuanshimingcheng_4_36.text
biaotou4_36 = "此次" + name_4_36 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_36)
print('账号为：',name_4_36 + " ；上传量：",shangchuan_4_36 , ' ；播放量：' , bofang_4_36 )

worksheet.write('H38', name_4_36,geshi)
worksheet.write('I38', shangchuan_4_36,geshi)
worksheet.write('J38', bofang_4_36,geshi)

# #观天
id_4_37 = '15961687741440'
url_4_37 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1294827&columnId=' + id_4_37
Browser.get(url_4_37)
time.sleep(4)

bf_4_37 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_37.text[-1].strip() == '万':
    bofang_4_37 = int(float(bf_4_37.text[:-1])*10000)
else:
    bofang_4_37 = bf_4_37.text

sc_4_37 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_37.text[-1].strip() == '集':
    shangchuan_4_37 = int(float(sc_4_37.text[:-1]))
else:
    shangchuan_4_37 = sc_4_37.text

yuanshimingcheng_4_37 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_37 = yuanshimingcheng_4_37.text
biaotou4_37 = "此次" + name_4_37 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_37)
print('账号为：',name_4_37 + " ；上传量：",shangchuan_4_37 , ' ；播放量：' , bofang_4_37 )

worksheet.write('H39', name_4_37,geshi)
worksheet.write('I39', shangchuan_4_37,geshi)
worksheet.write('J39', bofang_4_37,geshi)

# 摘点啥
id_4_38 = '15919485201930'
url_4_38 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1272365&columnId=' + id_4_38
Browser.get(url_4_38)
time.sleep(4)
Browser.refresh()
time.sleep(4)
bf_4_38 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_38.text[-1].strip() == '万':
    bofang_4_38 = int(float(bf_4_38.text[:-1])*10000)
else:
    bofang_4_38 = bf_4_38.text

sc_4_38 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_38.text[-1].strip() == '集':
    shangchuan_4_38 = int(float(sc_4_38.text[:-1]))
else:
    shangchuan_4_38 = sc_4_38.text

yuanshimingcheng_4_38 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_38 = yuanshimingcheng_4_38.text
biaotou4_38 = "此次" + name_4_38 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_38)
print('账号为：',name_4_38 + " ；上传量：",shangchuan_4_38 , ' ；播放量：' , bofang_4_38 )

worksheet.write('H40', name_4_38,geshi)
worksheet.write('I40', shangchuan_4_38,geshi)
worksheet.write('J40', bofang_4_38,geshi)

#暑期安全伴我行
id_4_39 = '15947090694820'
url_4_39 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1274114&columnId=' + id_4_39
Browser.get(url_4_39)
time.sleep(4)

bf_4_39 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_39.text[-1].strip() == '万':
    bofang_4_39 = int(float(bf_4_39.text[:-1])*10000)
else:
    bofang_4_39 = bf_4_39.text

sc_4_39 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_39.text[-1].strip() == '集':
    shangchuan_4_39 = int(float(sc_4_39.text[:-1]))
else:
    shangchuan_4_39 = sc_4_39.text

yuanshimingcheng_4_39 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_39 = yuanshimingcheng_4_39.text
biaotou4_39 = "此次" + name_4_39 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_39)
print('账号为：',name_4_39 + " ；上传量：",shangchuan_4_39 , ' ；播放量：' , bofang_4_39 )

worksheet.write('H41', name_4_39,geshi)
worksheet.write('I41', shangchuan_4_39,geshi)
worksheet.write('J41', bofang_4_39,geshi)

#乡村善治进行时
id_4_40 = '15901332396620'
url_4_40 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1211662&columnId=' + id_4_40
Browser.get(url_4_40)
time.sleep(4)

bf_4_40 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_40.text[-1].strip() == '万':
    bofang_4_40 = int(float(bf_4_40.text[:-1])*10000)
else:
    bofang_4_40 = bf_4_40.text

sc_4_40 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_40.text[-1].strip() == '集':
    shangchuan_4_40 = int(float(sc_4_40.text[:-1]))
else:
    shangchuan_4_40 = sc_4_40.text

yuanshimingcheng_4_40 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_40 = yuanshimingcheng_4_40.text
biaotou4_40 = "此次" + name_4_40 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_40)
print('账号为：',name_4_40 + " ；上传量：",shangchuan_4_40 , ' ；播放量：' , bofang_4_40 )

worksheet.write('H42', name_4_40,geshi)
worksheet.write('I42', shangchuan_4_40,geshi)
worksheet.write('J42', bofang_4_40,geshi)

#谣言粉碎机
id_4_41 = '15779380072320'
url_4_41 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1139805&columnId=' + id_4_41
Browser.get(url_4_41)
time.sleep(4)

bf_4_41 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_41.text[-1].strip() == '万':
    bofang_4_41 = int(float(bf_4_41.text[:-1])*10000)
else:
    bofang_4_41 = bf_4_41.text

sc_4_41 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_41.text[-1].strip() == '集':
    shangchuan_4_41 = int(float(sc_4_41.text[:-1]))
else:
    shangchuan_4_41 = sc_4_41.text

yuanshimingcheng_4_41 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_41 = yuanshimingcheng_4_41.text
biaotou4_41 = "此次" + name_4_41 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_41)
print('账号为：',name_4_41 + " ；上传量：",shangchuan_4_41 , ' ；播放量：' , bofang_4_41 )

worksheet.write('H43', name_4_41,geshi)
worksheet.write('I43', shangchuan_4_41,geshi)
worksheet.write('J43', bofang_4_41,geshi)

#高考后如何防骗
id_4_42 = '15957331245280'
url_4_42 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1286839&columnId=' + id_4_42
Browser.get(url_4_42)
time.sleep(4)

bf_4_42 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_42.text[-1].strip() == '万':
    bofang_4_42 = int(float(bf_4_42.text[:-1])*10000)
else:
    bofang_4_42 = bf_4_42.text

sc_4_42 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_42.text[-1].strip() == '集':
    shangchuan_4_42 = int(float(sc_4_42.text[:-1]))
else:
    shangchuan_4_42 = sc_4_42.text

yuanshimingcheng_4_42 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_42 = yuanshimingcheng_4_42.text
biaotou4_42 = "此次" + name_4_42 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_42)
print('账号为：',name_4_42 + " ；上传量：",shangchuan_4_42 , ' ；播放量：' , bofang_4_42 )

worksheet.write('H44', name_4_42,geshi)
worksheet.write('I44', shangchuan_4_42,geshi)
worksheet.write('J44', bofang_4_42,geshi)

#毛孩子的夏日护理
id_4_49 = '15967940605200'
url_4_49 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1296023&columnId=' + id_4_49
Browser.get(url_4_49)
time.sleep(4)

id_4_49 = '15967940605200'
url_4_49 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1296023&columnId=' + id_4_49
Browser.get(url_4_49)
time.sleep(4)

bf_4_49 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_49.text[-1].strip() == '万':
    bofang_4_49 = int(float(bf_4_49.text[:-1])*10000)
else:
    bofang_4_49 = bf_4_49.text

sc_4_49 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_49.text[-1].strip() == '集':
    shangchuan_4_49 = int(float(sc_4_49.text[:-1]))
else:
    shangchuan_4_49 = sc_4_49.text

yuanshimingcheng_4_49 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_49 = yuanshimingcheng_4_49.text
biaotou4_49 = "此次" + name_4_49 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_49)
print('账号为：',name_4_49 + " ；上传量：",shangchuan_4_49 , ' ；播放量：' , bofang_4_49 )

worksheet.write('H45', name_4_49,geshi)
worksheet.write('I45', shangchuan_4_49,geshi)
worksheet.write('J45', bofang_4_49,geshi)

#妈妈必备的育儿宝典
id_4_50 = '15971393100740'
url_4_50 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1301034&columnId=' + id_4_50
Browser.get(url_4_50)
time.sleep(4)

bf_4_50 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_50.text[-1].strip() == '万':
    bofang_4_50 = int(float(bf_4_50.text[:-1])*10000)
else:
    bofang_4_50 = bf_4_50.text

sc_4_50 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_50.text[-1].strip() == '集':
    shangchuan_4_50 = int(float(sc_4_50.text[:-1]))
else:
    shangchuan_4_50 = sc_4_50.text

yuanshimingcheng_4_50 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_50 = yuanshimingcheng_4_50.text
biaotou4_50 = "此次" + name_4_50 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_50)
print('账号为：',name_4_50 + " ；上传量：",shangchuan_4_50 , ' ；播放量：' , bofang_4_50 )

worksheet.write('H46', name_4_50,geshi)
worksheet.write('I46', shangchuan_4_50,geshi)
worksheet.write('J46', bofang_4_50,geshi)

#蓝朋友消防安全小课堂
id_4_51 = '15976574651080'
url_4_51 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1301287&columnId=' + id_4_51
Browser.get(url_4_51)
time.sleep(4)

bf_4_51 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_51.text[-1].strip() == '万':
    bofang_4_51 = int(float(bf_4_51.text[:-1])*10000)
else:
    bofang_4_51 = bf_4_51.text

sc_4_51 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_51.text[-1].strip() == '集':
    shangchuan_4_51 = int(float(sc_4_51.text[:-1]))
else:
    shangchuan_4_51 = sc_4_51.text

yuanshimingcheng_4_51 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_51 = yuanshimingcheng_4_51.text
biaotou4_51 = "此次" + name_4_51 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_51)
print('账号为：',name_4_51 + " ；上传量：",shangchuan_4_51 , ' ；播放量：' , bofang_4_51 )

worksheet.write('H47', name_4_51,geshi)
worksheet.write('I47', shangchuan_4_51,geshi)
worksheet.write('J47', bofang_4_51,geshi)

#干杯快乐
id_4_52 = '15973095473070'
url_4_52 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1298957&columnId=' + id_4_52
Browser.get(url_4_52)
time.sleep(4)

bf_4_52 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_52.text[-1].strip() == '万':
    bofang_4_52 = int(float(bf_4_52.text[:-1])*10000)
else:
    bofang_4_52 = bf_4_52.text

sc_4_52 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_52.text[-1].strip() == '集':
    shangchuan_4_52 = int(float(sc_4_52.text[:-1]))
else:
    shangchuan_4_52 = sc_4_52.text

yuanshimingcheng_4_52 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_52 = yuanshimingcheng_4_52.text
biaotou4_52 = "此次" + name_4_52 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_52)
print('账号为：',name_4_52 + " ；上传量：",shangchuan_4_52 , ' ；播放量：' , bofang_4_52 )

worksheet.write('H48', name_4_52,geshi)
worksheet.write('I48', shangchuan_4_52,geshi)
worksheet.write('J48', bofang_4_52,geshi)

#我的医生朋友
id_4_53 = '15977152774090'
url_4_53 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1300966&columnId=' + id_4_53
Browser.get(url_4_53)
time.sleep(4)

bf_4_53 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_53.text[-1].strip() == '万':
    bofang_4_53 = int(float(bf_4_53.text[:-1])*10000)
else:
    bofang_4_53 = bf_4_53.text

sc_4_53 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_53.text[-1].strip() == '集':
    shangchuan_4_53 = int(float(sc_4_53.text[:-1]))
else:
    shangchuan_4_53 = sc_4_53.text

yuanshimingcheng_4_53 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_53 = yuanshimingcheng_4_53.text
biaotou4_53 = "此次" + name_4_53 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_53)
print('账号为：',name_4_53 + " ；上传量：",shangchuan_4_53 , ' ；播放量：' , bofang_4_53 )

worksheet.write('H49', name_4_53,geshi)
worksheet.write('I49', shangchuan_4_53,geshi)
worksheet.write('J49', bofang_4_53,geshi)

#神奇的动物
id_4_54 = '15977249761290'
url_4_54 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1304245&columnId=' + id_4_54
Browser.get(url_4_54)
time.sleep(4)

bf_4_54 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_54.text[-1].strip() == '万':
    bofang_4_54 = int(float(bf_4_54.text[:-1])*10000)
else:
    bofang_4_54 = bf_4_54.text

sc_4_54 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_54.text[-1].strip() == '集':
    shangchuan_4_54 = int(float(sc_4_54.text[:-1]))
else:
    shangchuan_4_54 = sc_4_54.text

yuanshimingcheng_4_54 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_54 = yuanshimingcheng_4_54.text
biaotou4_54 = "此次" + name_4_54 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_54)
print('账号为：',name_4_54 + " ；上传量：",shangchuan_4_54 , ' ；播放量：' , bofang_4_54 )

worksheet.write('H50', name_4_54,geshi)
worksheet.write('I50', shangchuan_4_54,geshi)
worksheet.write('J50', bofang_4_54,geshi)

#扶贫日记
id_4_55 = '15984062550340'
url_4_55 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1305075&columnId=' + id_4_55
Browser.get(url_4_55)
time.sleep(4)

bf_4_55 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_55.text[-1].strip() == '万':
    bofang_4_55 = int(float(bf_4_55.text[:-1])*10000)
else:
    bofang_4_55 = bf_4_55.text

sc_4_55 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_55.text[-1].strip() == '集':
    shangchuan_4_55 = int(float(sc_4_55.text[:-1]))
else:
    shangchuan_4_55 = sc_4_55.text

yuanshimingcheng_4_55 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_55 = yuanshimingcheng_4_55.text
biaotou4_55 = "此次" + name_4_55 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_55)
print('账号为：',name_4_55 + " ；上传量：",shangchuan_4_55 , ' ；播放量：' , bofang_4_55 )

worksheet.write('H51', name_4_55,geshi)
worksheet.write('I51', shangchuan_4_55,geshi)
worksheet.write('J51', bofang_4_55,geshi)

#戏曲中的爱情故事
id_4_56 = '15983407706590'
url_4_56 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1304950&columnId=' + id_4_56
Browser.get(url_4_56)
time.sleep(4)

bf_4_56 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_56.text[-1].strip() == '万':
    bofang_4_56 = int(float(bf_4_56.text[:-1])*10000)
else:
    bofang_4_56 = bf_4_56.text

sc_4_56 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_56.text[-1].strip() == '集':
    shangchuan_4_56 = int(float(sc_4_56.text[:-1]))
else:
    shangchuan_4_56 = sc_4_56.text

yuanshimingcheng_4_56 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_56 = yuanshimingcheng_4_56.text
biaotou4_56 = "此次" + name_4_56 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_56)
print('账号为：',name_4_56 + " ；上传量：",shangchuan_4_56 , ' ；播放量：' , bofang_4_56 )

worksheet.write('H52', name_4_56,geshi)
worksheet.write('I52', shangchuan_4_56,geshi)
worksheet.write('J52', bofang_4_56,geshi)

#乡村热搜榜
id_4_57 = '15979154548250'
url_4_57 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1304521&columnId=' + id_4_57
Browser.get(url_4_57)
time.sleep(4)

bf_4_57 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_57.text[-1].strip() == '万':
    bofang_4_57 = int(float(bf_4_57.text[:-1])*10000)
else:
    bofang_4_57 = bf_4_57.text

sc_4_57 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_57.text[-1].strip() == '集':
    shangchuan_4_57 = int(float(sc_4_57.text[:-1]))
else:
    shangchuan_4_57 = sc_4_57.text

yuanshimingcheng_4_57 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_57 = yuanshimingcheng_4_57.text
biaotou4_57 = "此次" + name_4_57 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_57)
print('账号为：',name_4_57 + " ；上传量：",shangchuan_4_57 , ' ；播放量：' , bofang_4_57 )

worksheet.write('H53', name_4_57,geshi)
worksheet.write('I53', shangchuan_4_57,geshi)
worksheet.write('J53', bofang_4_57,geshi)


#说出你的秘密
id_4_58 = '15994431153040'
url_4_58 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1310058&columnId=' + id_4_58
Browser.get(url_4_58)
time.sleep(4)

bf_4_58 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_58.text[-1].strip() == '万':
    bofang_4_58 = int(float(bf_4_58.text[:-1])*10000)
else:
    bofang_4_58 = bf_4_58.text

sc_4_58 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_58.text[-1].strip() == '集':
    shangchuan_4_58 = int(float(sc_4_58.text[:-1]))
else:
    shangchuan_4_58 = sc_4_58.text

yuanshimingcheng_4_58 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_58 = yuanshimingcheng_4_58.text
biaotou4_58 = "此次" + name_4_58 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_58)
print('账号为：',name_4_58 + " ；上传量：",shangchuan_4_58 , ' ；播放量：' , bofang_4_58 )

worksheet.write('H54', name_4_58,geshi)
worksheet.write('I54', shangchuan_4_58,geshi)
worksheet.write('J54', bofang_4_58,geshi)

#谈法不伤感情
id_4_59 = '15995583914050'
url_4_59 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1311635&columnId=' + id_4_59
Browser.get(url_4_59)
time.sleep(4)

bf_4_59 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_59.text[-1].strip() == '万':
    bofang_4_59 = int(float(bf_4_59.text[:-1])*10000)
else:
    bofang_4_59 = bf_4_59.text

sc_4_59 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_59.text[-1].strip() == '集':
    shangchuan_4_59 = int(float(sc_4_59.text[:-1]))
else:
    shangchuan_4_59 = sc_4_59.text

yuanshimingcheng_4_59 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_59 = yuanshimingcheng_4_59.text
biaotou4_59 = "此次" + name_4_59 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_59)
print('账号为：',name_4_59 + " ；上传量：",shangchuan_4_59 , ' ；播放量：' , bofang_4_59 )

worksheet.write('H55', name_4_59,geshi)
worksheet.write('I55', shangchuan_4_59,geshi)
worksheet.write('J55', bofang_4_59,geshi)

#必吃研究所
id_4_60 = '15996405779040'
url_4_60 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1312658&columnId=' + id_4_60
Browser.get(url_4_60)
time.sleep(4)

bf_4_60 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_60.text[-1].strip() == '万':
    bofang_4_60 = int(float(bf_4_60.text[:-1])*10000)
else:
    bofang_4_60 = bf_4_60.text

sc_4_60 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_60.text[-1].strip() == '集':
    shangchuan_4_60 = int(float(sc_4_60.text[:-1]))
else:
    shangchuan_4_60 = sc_4_60.text

yuanshimingcheng_4_60 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_60 = yuanshimingcheng_4_60.text
biaotou4_60 = "此次" + name_4_60 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_60)
print('账号为：',name_4_60 + " ；上传量：",shangchuan_4_60 , ' ；播放量：' , bofang_4_60 )

worksheet.write('H56', name_4_60,geshi)
worksheet.write('I56', shangchuan_4_60,geshi)
worksheet.write('J56', bofang_4_60,geshi)

#新鲜说法
id_4_61 = '16006737562690'
url_4_61 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1319614&columnId=' + id_4_61
Browser.get(url_4_61)
time.sleep(4)

bf_4_61 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_61.text[-1].strip() == '万':
    bofang_4_61 = int(float(bf_4_61.text[:-1])*10000)
else:
    bofang_4_61 = bf_4_61.text

sc_4_61 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_61.text[-1].strip() == '集':
    shangchuan_4_61 = int(float(sc_4_61.text[:-1]))
else:
    shangchuan_4_61 = sc_4_61.text

yuanshimingcheng_4_61 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_61 = yuanshimingcheng_4_61.text
biaotou4_61 = "此次" + name_4_61 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_61)
print('账号为：',name_4_61 + " ；上传量：",shangchuan_4_61 , ' ；播放量：' , bofang_4_61 )

worksheet.write('H57', name_4_61,geshi)
worksheet.write('I57', shangchuan_4_61,geshi)
worksheet.write('J57', bofang_4_61,geshi)

#两分钟拯救法律小白
id_4_62 = '16007042575020'
url_4_62 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1319917&columnId=' + id_4_62
Browser.get(url_4_62)
time.sleep(4)

bf_4_62 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_62.text[-1].strip() == '万':
    bofang_4_62 = int(float(bf_4_62.text[:-1])*10000)
else:
    bofang_4_62 = bf_4_62.text

sc_4_62 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_62.text[-1].strip() == '集':
    shangchuan_4_62 = int(float(sc_4_62.text[:-1]))
else:
    shangchuan_4_62 = sc_4_62.text

yuanshimingcheng_4_62 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_62 = yuanshimingcheng_4_62.text
biaotou4_62 = "此次" + name_4_62 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_62)
print('账号为：',name_4_62 + " ；上传量：",shangchuan_4_62 , ' ；播放量：' , bofang_4_62 )

worksheet.write('H58', name_4_62,geshi)
worksheet.write('I58', shangchuan_4_62,geshi)
worksheet.write('J58', bofang_4_62,geshi)

#这个案子不简单
id_4_63 = '16006672904940'
url_4_63 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1319818&columnId=' + id_4_63
Browser.get(url_4_63)
time.sleep(4)

bf_4_63 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_63.text[-1].strip() == '万':
    bofang_4_63 = int(float(bf_4_63.text[:-1])*10000)
else:
    bofang_4_63 = bf_4_63.text

sc_4_63 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_63.text[-1].strip() == '集':
    shangchuan_4_63 = int(float(sc_4_63.text[:-1]))
else:
    shangchuan_4_63 = sc_4_63.text

yuanshimingcheng_4_63 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_63 = yuanshimingcheng_4_63.text
biaotou4_63 = "此次" + name_4_63 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_63)
print('账号为：',name_4_63 + " ；上传量：",shangchuan_4_63 , ' ；播放量：' , bofang_4_63 )

worksheet.write('H59', name_4_63,geshi)
worksheet.write('I59', shangchuan_4_63,geshi)
worksheet.write('J59', bofang_4_63,geshi)

#医院里的故事
id_4_64 = '16006589988440'
url_4_64 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1319244&columnId=' + id_4_64
Browser.get(url_4_64)
time.sleep(4)

bf_4_64 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_64.text[-1].strip() == '万':
    bofang_4_64 = int(float(bf_4_64.text[:-1])*10000)
else:
    bofang_4_64 = bf_4_64.text

sc_4_64 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_64.text[-1].strip() == '集':
    shangchuan_4_64 = int(float(sc_4_64.text[:-1]))
else:
    shangchuan_4_64 = sc_4_64.text

yuanshimingcheng_4_64 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_64 = yuanshimingcheng_4_64.text
biaotou4_64 = "此次" + name_4_64 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_64)
print('账号为：',name_4_64 + " ；上传量：",shangchuan_4_64 , ' ；播放量：' , bofang_4_64 )

worksheet.write('H60', name_4_64,geshi)
worksheet.write('I60', shangchuan_4_64,geshi)
worksheet.write('J60', bofang_4_64,geshi)

#大医生上线了
id_4_65 = '16006864079190'
url_4_65 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1319789&columnId=' + id_4_65
Browser.get(url_4_65)
time.sleep(4)

bf_4_65 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_65.text[-1].strip() == '万':
    bofang_4_65 = int(float(bf_4_65.text[:-1])*10000)
else:
    bofang_4_65 = bf_4_65.text

sc_4_65 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_65.text[-1].strip() == '集':
    shangchuan_4_65 = int(float(sc_4_65.text[:-1]))
else:
    shangchuan_4_65 = sc_4_65.text

yuanshimingcheng_4_65 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_65 = yuanshimingcheng_4_65.text
biaotou4_65 = "此次" + name_4_65 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_65)
print('账号为：',name_4_65 + " ；上传量：",shangchuan_4_65 , ' ；播放量：' , bofang_4_65 )

worksheet.write('H61', name_4_65,geshi)
worksheet.write('I61', shangchuan_4_65,geshi)
worksheet.write('J61', bofang_4_65,geshi)

#健康捉妖记
id_4_66 = '16006745781810'
url_4_66 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1320619&columnId=' + id_4_66
Browser.get(url_4_66)
time.sleep(4)

bf_4_66 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_66.text[-1].strip() == '万':
    bofang_4_66 = int(float(bf_4_66.text[:-1])*10000)
else:
    bofang_4_66 = bf_4_66.text

sc_4_66 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_66.text[-1].strip() == '集':
    shangchuan_4_66 = int(float(sc_4_66.text[:-1]))
else:
    shangchuan_4_66 = sc_4_66.text

yuanshimingcheng_4_66 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_66 = yuanshimingcheng_4_66.text
biaotou4_66 = "此次" + name_4_66 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_66)
print('账号为：',name_4_66 + " ；上传量：",shangchuan_4_66 , ' ；播放量：' , bofang_4_66 )

worksheet.write('H62', name_4_66,geshi)
worksheet.write('I62', shangchuan_4_66,geshi)
worksheet.write('J62', bofang_4_66,geshi)

#养生妙招听我说
id_4_67 = '16007026364900'
url_4_67 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1319896&columnId=' + id_4_67
Browser.get(url_4_67)
time.sleep(4)

bf_4_67 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_67.text[-1].strip() == '万':
    bofang_4_67 = int(float(bf_4_67.text[:-1])*10000)
else:
    bofang_4_67 = bf_4_67.text

sc_4_67 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_67.text[-1].strip() == '集':
    shangchuan_4_67 = int(float(sc_4_67.text[:-1]))
else:
    shangchuan_4_67 = sc_4_67.text

yuanshimingcheng_4_67 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_67 = yuanshimingcheng_4_67.text
biaotou4_67 = "此次" + name_4_67 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_67)
print('账号为：',name_4_67 + " ；上传量：",shangchuan_4_67 , ' ；播放量：' , bofang_4_67 )

worksheet.write('H63', name_4_67,geshi)
worksheet.write('I63', shangchuan_4_67,geshi)
worksheet.write('J63', bofang_4_67,geshi)

#玥吃玥开心
id_4_68 = '16006641362910'
url_4_68 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1319466&columnId=' + id_4_68
Browser.get(url_4_68)
time.sleep(4)

bf_4_68 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_68.text[-1].strip() == '万':
    bofang_4_68 = int(float(bf_4_68.text[:-1])*10000)
else:
    bofang_4_68 = bf_4_68.text

sc_4_68 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_68.text[-1].strip() == '集':
    shangchuan_4_68 = int(float(sc_4_68.text[:-1]))
else:
    shangchuan_4_68 = sc_4_68.text

yuanshimingcheng_4_68 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_68 = yuanshimingcheng_4_68.text
biaotou4_68 = "此次" + name_4_68 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_68)
print('账号为：',name_4_68 + " ；上传量：",shangchuan_4_68 , ' ；播放量：' , bofang_4_68 )

worksheet.write('H64', name_4_68,geshi)
worksheet.write('I64', shangchuan_4_68,geshi)
worksheet.write('J64', bofang_4_68,geshi)

#丰收BGM
id_4_69 = '16007028842730'
url_4_69 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1319897&columnId=' + id_4_69
Browser.get(url_4_69)
time.sleep(4)

bf_4_69 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_69.text[-1].strip() == '万':
    bofang_4_69 = int(float(bf_4_69.text[:-1])*10000)
else:
    bofang_4_69 = bf_4_69.text

sc_4_69 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_69.text[-1].strip() == '集':
    shangchuan_4_69 = int(float(sc_4_69.text[:-1]))
else:
    shangchuan_4_69 = sc_4_69.text

yuanshimingcheng_4_69 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_69 = yuanshimingcheng_4_69.text
biaotou4_69 = "此次" + name_4_69 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_69)
print('账号为：',name_4_69 + " ；上传量：",shangchuan_4_69 , ' ；播放量：' , bofang_4_69 )

worksheet.write('H65', name_4_69,geshi)
worksheet.write('I65', shangchuan_4_69,geshi)
worksheet.write('J65', bofang_4_69,geshi)

#梨园乡韵
id_4_70 = '16006901273370'
url_4_70 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1319847&columnId=' + id_4_70
Browser.get(url_4_70)
time.sleep(4)

bf_4_70 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_70.text[-1].strip() == '万':
    bofang_4_70 = int(float(bf_4_70.text[:-1])*10000)
else:
    bofang_4_70 = bf_4_70.text

sc_4_70 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_70.text[-1].strip() == '集':
    shangchuan_4_70 = int(float(sc_4_70.text[:-1]))
else:
    shangchuan_4_70 = sc_4_70.text

yuanshimingcheng_4_70 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_70 = yuanshimingcheng_4_70.text
biaotou4_70 = "此次" + name_4_70 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_70)
print('账号为：',name_4_70 + " ；上传量：",shangchuan_4_70 , ' ；播放量：' , bofang_4_70 )

worksheet.write('H66', name_4_70,geshi)
worksheet.write('I66', shangchuan_4_70,geshi)
worksheet.write('J66', bofang_4_70,geshi)

#山水乡愁
id_4_71 = '16006741037570'
url_4_71 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1319575&columnId=' + id_4_71
Browser.get(url_4_71)
time.sleep(4)

bf_4_71 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_71.text[-1].strip() == '万':
    bofang_4_71 = int(float(bf_4_71.text[:-1])*10000)
else:
    bofang_4_71 = bf_4_71.text

sc_4_71 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_71.text[-1].strip() == '集':
    shangchuan_4_71 = int(float(sc_4_71.text[:-1]))
else:
    shangchuan_4_71 = sc_4_71.text

yuanshimingcheng_4_71 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_71 = yuanshimingcheng_4_71.text
biaotou4_71 = "此次" + name_4_71 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_71)
print('账号为：',name_4_71 + " ；上传量：",shangchuan_4_71 , ' ；播放量：' , bofang_4_71 )

worksheet.write('H67', name_4_71,geshi)
worksheet.write('I67', shangchuan_4_71,geshi)
worksheet.write('J67', bofang_4_71,geshi)

#乡村音乐
id_4_72 = '16006866713650'
url_4_72 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1319792&columnId=' + id_4_72
Browser.get(url_4_72)
time.sleep(4)

bf_4_72 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_72.text[-1].strip() == '万':
    bofang_4_72 = int(float(bf_4_72.text[:-1])*10000)
else:
    bofang_4_72 = bf_4_72.text

sc_4_72 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_72.text[-1].strip() == '集':
    shangchuan_4_72 = int(float(sc_4_72.text[:-1]))
else:
    shangchuan_4_72 = sc_4_72.text

yuanshimingcheng_4_72 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_72 = yuanshimingcheng_4_72.text
biaotou4_72 = "此次" + name_4_72 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_72)
print('账号为：',name_4_72 + " ；上传量：",shangchuan_4_72 , ' ；播放量：' , bofang_4_72 )

worksheet.write('H68', name_4_72,geshi)
worksheet.write('I68', shangchuan_4_72,geshi)
worksheet.write('J68', bofang_4_72,geshi)

#田间笑语
id_4_73 = '16006870768350'
url_4_73 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1438639&columnId=' + id_4_73
Browser.get(url_4_73)
time.sleep(4)

bf_4_73 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_73.text[-1].strip() == '万':
    bofang_4_73 = int(float(bf_4_73.text[:-1])*10000)
else:
    bofang_4_73 = bf_4_73.text

sc_4_73 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_73.text[-1].strip() == '集':
    shangchuan_4_73 = int(float(sc_4_73.text[:-1]))
else:
    shangchuan_4_73 = sc_4_73.text

yuanshimingcheng_4_73 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_73 = yuanshimingcheng_4_73.text
biaotou4_73 = "此次" + name_4_73 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_73)
print('账号为：',name_4_73 + " ；上传量：",shangchuan_4_73 , ' ；播放量：' , bofang_4_73 )

worksheet.write('H69', name_4_73,geshi)
worksheet.write('I69', shangchuan_4_73,geshi)
worksheet.write('J69', bofang_4_73,geshi)

#农主播朋友圈
id_4_83 = '16007045664830'
url_4_83 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1319921&columnId=' + id_4_83
Browser.get(url_4_83)
time.sleep(4)

bf_4_83 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_83.text[-1].strip() == '万':
    bofang_4_83 = int(float(bf_4_83.text[:-1])*10000)
else:
    bofang_4_83 = bf_4_83.text

sc_4_83 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_83.text[-1].strip() == '集':
    shangchuan_4_83 = int(float(sc_4_83.text[:-1]))
else:
    shangchuan_4_83 = sc_4_83.text

yuanshimingcheng_4_83 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_83 = yuanshimingcheng_4_83.text
biaotou4_83 = "此次" + name_4_83 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_83)
print('账号为：',name_4_83 + " ；上传量：",shangchuan_4_83 , ' ；播放量：' , bofang_4_83 )

worksheet.write('H70', name_4_83,geshi)
worksheet.write('I70', shangchuan_4_83,geshi)
worksheet.write('J70', bofang_4_83,geshi)

#田园频道来啦
id_4_84 = '16006666249940'
url_4_84 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1319394&columnId=' + id_4_84
Browser.get(url_4_84)
time.sleep(4)

bf_4_84 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_84.text[-1].strip() == '万':
    bofang_4_84 = int(float(bf_4_84.text[:-1])*10000)
else:
    bofang_4_84 = bf_4_84.text

sc_4_84 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_84.text[-1].strip() == '集':
    shangchuan_4_84 = int(float(sc_4_84.text[:-1]))
else:
    shangchuan_4_84 = sc_4_84.text

yuanshimingcheng_4_84 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_84 = yuanshimingcheng_4_84.text
biaotou4_84 = "此次" + name_4_84 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_84)
print('账号为：',name_4_84 + " ；上传量：",shangchuan_4_84 , ' ；播放量：' , bofang_4_84 )

worksheet.write('H71', name_4_84,geshi)
worksheet.write('I71', shangchuan_4_84,geshi)
worksheet.write('J71', bofang_4_84,geshi)

#双十一|“剁手”没那么简单
id_4_85 = '16040628241380'
url_4_85 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1341985&columnId=' + id_4_85
Browser.get(url_4_85)
time.sleep(4)

bf_4_85 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_85.text[-1].strip() == '万':
    bofang_4_85 = int(float(bf_4_85.text[:-1])*10000)
else:
    bofang_4_85 = bf_4_85.text

sc_4_85 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_85.text[-1].strip() == '集':
    shangchuan_4_85 = int(float(sc_4_85.text[:-1]))
else:
    shangchuan_4_85 = sc_4_85.text

yuanshimingcheng_4_85 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_85 = yuanshimingcheng_4_85.text
biaotou4_85 = "此次" + name_4_85 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_85)
print('账号为：',name_4_85 + " ；上传量：",shangchuan_4_85 , ' ；播放量：' , bofang_4_85 )

worksheet.write('H72', name_4_85,geshi)
worksheet.write('I72', shangchuan_4_85,geshi)
worksheet.write('J72', bofang_4_85,geshi)

#律师爸爸说法律
id_4_86 = '16060960598850'
url_4_86 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1364856&columnId=' + id_4_86
Browser.get(url_4_86)
time.sleep(4)

bf_4_86 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_86.text[-1].strip() == '万':
    bofang_4_86 = int(float(bf_4_86.text[:-1])*10000)
else:
    bofang_4_86 = bf_4_86.text

sc_4_86 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_86.text[-1].strip() == '集':
    shangchuan_4_86 = int(float(sc_4_86.text[:-1]))
else:
    shangchuan_4_86 = sc_4_86.text

yuanshimingcheng_4_86 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_86 = yuanshimingcheng_4_86.text
biaotou4_86 = "此次" + name_4_86 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_86)
print('账号为：',name_4_86 + " ；上传量：",shangchuan_4_86 , ' ；播放量：' , bofang_4_86 )

worksheet.write('H73', name_4_86,geshi)
worksheet.write('I73', shangchuan_4_86,geshi)
worksheet.write('J73', bofang_4_86,geshi)

# #暖暖的灶台（特别节目）
# id_4_88 = '16093839231540'
# url_4_88 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1390482&columnId=' + id_4_88
# Browser.get(url_4_88)
# time.sleep(4)
#
# bf_4_88 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
# if bf_4_88.text[-1].strip() == '万':
#     bofang_4_88 = int(float(bf_4_88.text[:-1])*10000)
# else:
#     bofang_4_88 = bf_4_88.text
#
# sc_4_88 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
# if sc_4_88.text[-1].strip() == '集':
#     shangchuan_4_88 = int(float(sc_4_88.text[:-1]))
# else:
#     shangchuan_4_88 = sc_4_88.text
#
# yuanshimingcheng_4_88 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
# name_4_88 = yuanshimingcheng_4_88.text
# biaotou4_88 = "此次" + name_4_88 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))
#
# print (biaotou4_88)
# print('账号为：',name_4_88 + " ；上传量：",shangchuan_4_88 , ' ；播放量：' , bofang_4_88 )
#
# worksheet.write('H74', name_4_88,geshi)
# worksheet.write('I74', shangchuan_4_88,geshi)
# worksheet.write('J74', bofang_4_88,geshi)

#“熊孩子”的安全攻略
id_4_89 = '16115637391070'
url_4_89 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1406291&columnId=' + id_4_89
Browser.get(url_4_89)
time.sleep(4)

bf_4_89 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_89.text[-1].strip() == '万':
    bofang_4_89 = int(float(bf_4_89.text[:-1])*10000)
else:
    bofang_4_89 = bf_4_89.text

sc_4_89 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_89.text[-1].strip() == '集':
    shangchuan_4_89 = int(float(sc_4_89.text[:-1]))
else:
    shangchuan_4_89 = sc_4_89.text

yuanshimingcheng_4_89 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_89 = yuanshimingcheng_4_89.text
biaotou4_89 = "此次" + name_4_89 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_89)
print('账号为：',name_4_89 + " ；上传量：",shangchuan_4_89 , ' ；播放量：' , bofang_4_89 )

worksheet.write('H75', name_4_89,geshi)
worksheet.write('I75', shangchuan_4_89,geshi)
worksheet.write('J75', bofang_4_89,geshi)

# #袁隆平采访实录
id_4_90 = '16121451546140'
url_4_90 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1410366&columnId=' + id_4_90
Browser.get(url_4_90)
time.sleep(4)

bf_4_90 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_90.text[-1].strip() == '万':
    bofang_4_90 = int(float(bf_4_90.text[:-1])*10000)
else:
    bofang_4_90 = bf_4_90.text

sc_4_90 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_90.text[-1].strip() == '集':
    shangchuan_4_90 = int(float(sc_4_90.text[:-1]))
else:
    shangchuan_4_90 = sc_4_90.text

yuanshimingcheng_4_90 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_90 = yuanshimingcheng_4_90.text
biaotou4_90 = "此次" + name_4_90 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_90)
print('账号为：',name_4_90 + " ；上传量：",shangchuan_4_90 , ' ；播放量：' , bofang_4_90 )

worksheet.write('H76', name_4_90,geshi)
worksheet.write('I76', shangchuan_4_90,geshi)
worksheet.write('J76', bofang_4_90,geshi)

#非常青年研究所
id_4_91 = '16120733979980'
url_4_91 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1412576&columnId=' + id_4_91
Browser.get(url_4_91)
time.sleep(4)

bf_4_91 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_91.text[-1].strip() == '万':
    bofang_4_91 = int(float(bf_4_91.text[:-1])*10000)
else:
    bofang_4_91 = bf_4_91.text

sc_4_91 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_91.text[-1].strip() == '集':
    shangchuan_4_91 = int(float(sc_4_91.text[:-1]))
else:
    shangchuan_4_91 = sc_4_91.text

yuanshimingcheng_4_91 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_91 = yuanshimingcheng_4_91.text
biaotou4_91 = "此次" + name_4_91 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_91)
print('账号为：',name_4_91 + " ；上传量：",shangchuan_4_91 , ' ；播放量：' , bofang_4_91 )

worksheet.write('H77', name_4_91,geshi)
worksheet.write('I77', shangchuan_4_91,geshi)
worksheet.write('J77', bofang_4_91,geshi)

#你好，童年
id_4_92 = '16121427028420'
url_4_92 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1410294&columnId=' + id_4_92
Browser.get(url_4_92)
time.sleep(4)

bf_4_92 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_92.text[-1].strip() == '万':
    bofang_4_92 = int(float(bf_4_92.text[:-1])*10000)
else:
    bofang_4_92 = bf_4_92.text

sc_4_92 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_92.text[-1].strip() == '集':
    shangchuan_4_92 = int(float(sc_4_92.text[:-1]))
else:
    shangchuan_4_92 = sc_4_92.text

yuanshimingcheng_4_92 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_92 = yuanshimingcheng_4_92.text
biaotou4_92 = "此次" + name_4_92 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_92)
print('账号为：',name_4_92 + " ；上传量：",shangchuan_4_92 , ' ；播放量：' , bofang_4_92 )

worksheet.write('H78', name_4_92,geshi)
worksheet.write('I78', shangchuan_4_92,geshi)
worksheet.write('J78', bofang_4_92,geshi)

#开麦吧，原野人
id_4_93 = '16124037936500'
url_4_93 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1412956&columnId=' + id_4_93
Browser.get(url_4_93)
time.sleep(4)

bf_4_93 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_93.text[-1].strip() == '万':
    bofang_4_93 = int(float(bf_4_93.text[:-1])*10000)
else:
    bofang_4_93 = bf_4_93.text

sc_4_93 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_93.text[-1].strip() == '集':
    shangchuan_4_93 = int(float(sc_4_93.text[:-1]))
else:
    shangchuan_4_93 = sc_4_93.text

yuanshimingcheng_4_93 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_93 = yuanshimingcheng_4_93.text
biaotou4_93 = "此次" + name_4_93 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_93)
print('账号为：',name_4_93 + " ；上传量：",shangchuan_4_93 , ' ；播放量：' , bofang_4_93 )

worksheet.write('H79', name_4_93,geshi)
worksheet.write('I79', shangchuan_4_93,geshi)
worksheet.write('J79', bofang_4_93,geshi)

#年俗探秘
id_4_94 = '16128907749210'
url_4_94 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1418041&columnId=' + id_4_94
Browser.get(url_4_94)
time.sleep(4)

bf_4_94 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_94.text[-1].strip() == '万':
    bofang_4_94 = int(float(bf_4_94.text[:-1])*10000)
else:
    bofang_4_94 = bf_4_94.text

sc_4_94 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_94.text[-1].strip() == '集':
    shangchuan_4_94 = int(float(sc_4_94.text[:-1]))
else:
    shangchuan_4_94 = sc_4_94.text

yuanshimingcheng_4_94 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_94 = yuanshimingcheng_4_94.text
biaotou4_94 = "此次" + name_4_94 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_94)
print('账号为：',name_4_94 + " ；上传量：",shangchuan_4_94 , ' ；播放量：' , bofang_4_94 )

worksheet.write('H80', name_4_94,geshi)
worksheet.write('I80', shangchuan_4_94,geshi)
worksheet.write('J80', bofang_4_94,geshi)

#中国三农报道(云听)
id_4_96 = '16104126182600'
url_4_96 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1405327&columnId=' + id_4_96
Browser.get(url_4_96)
time.sleep(4)

bf_4_96 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_96.text[-1].strip() == '万':
    bofang_4_96 = int(float(bf_4_96.text[:-1])*10000)
else:
    bofang_4_96 = bf_4_96.text

sc_4_96 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_96.text[-1].strip() == '集':
    shangchuan_4_96 = int(float(sc_4_96.text[:-1]))
else:
    shangchuan_4_96 = sc_4_96.text

yuanshimingcheng_4_96 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_96 = yuanshimingcheng_4_96.text
biaotou4_96 = "此次" + name_4_96 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_96)
print('账号为：',name_4_96 + " ；上传量：",shangchuan_4_96 , ' ；播放量：' , bofang_4_96 )

worksheet.write('H81', name_4_96,geshi)
worksheet.write('I81', shangchuan_4_96,geshi)
worksheet.write('J81', bofang_4_96,geshi)

#让你脑洞大开的法律冷知识
id_4_82 = '16006778761100'
url_4_82 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1321020&columnId=' + id_4_82
Browser.get(url_4_82)
time.sleep(4)

bf_4_82 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_82.text[-1].strip() == '万':
    bofang_4_82 = int(float(bf_4_82.text[:-1])*10000)
else:
    bofang_4_82 = bf_4_82.text

sc_4_82 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_82.text[-1].strip() == '集':
    shangchuan_4_82 = int(float(sc_4_82.text[:-1]))
else:
    shangchuan_4_82 = sc_4_82.text

yuanshimingcheng_4_82 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_82 = yuanshimingcheng_4_82.text
biaotou4_82 = "此次" + name_4_82 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_82)
print('账号为：',name_4_82 + " ；上传量：",shangchuan_4_82 , ' ；播放量：' , bofang_4_82 )

worksheet.write('H82', name_4_82,geshi)
worksheet.write('I82', shangchuan_4_82,geshi)
worksheet.write('J82', bofang_4_82,geshi)

# 钟大厨的美食江湖
id_200_45 = '16048153614080'
url_200_45 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1460750&columnId=' + id_200_45
Browser.get(url_200_45)
time.sleep(4)

bf_200_45 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_45.text[-1].strip() == '万':
    bofang_200_45 = int(float(bf_200_45.text[:-1])*10000)
else:
    bofang_200_45 = bf_200_45.text

sc_200_45 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_45.text[-1].strip() == '集':
    shangchuan_200_45 = int(float(sc_200_45.text[:-1]))
else:
    shangchuan_200_45 = sc_200_45.text

yuanshimingcheng_200_45 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_45 = yuanshimingcheng_200_45.text
biaotou200_45 = "此次" + name_200_45 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_45)
print('账号为：',name_200_45 + " ；上传量：",shangchuan_200_45 , ' ；播放量：' , bofang_200_45 )

worksheet.write('H83', name_200_45,geshi)
worksheet.write('I83', shangchuan_200_45,geshi)
worksheet.write('J83', bofang_200_45,geshi)

#生命对话
id_4_99 = '16160475864240'
url_4_99 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1440296&columnId=' + id_4_99
Browser.get(url_4_99)
time.sleep(4)

bf_4_99 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_99.text[-1].strip() == '万':
    bofang_4_99 = int(float(bf_4_99.text[:-1])*10000)
else:
    bofang_4_99 = bf_4_99.text

sc_4_99 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_99.text[-1].strip() == '集':
    shangchuan_4_99 = int(float(sc_4_99.text[:-1]))
else:
    shangchuan_4_99 = sc_4_99.text

yuanshimingcheng_4_99 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_99 = yuanshimingcheng_4_99.text
biaotou4_99 = "此次" + name_4_99 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_99)
print('账号为：',name_4_99 + " ；上传量：",shangchuan_4_99 , ' ；播放量：' , bofang_4_99 )

worksheet.write('H84', name_4_99,geshi)
worksheet.write('I84', shangchuan_4_99,geshi)
worksheet.write('J84', bofang_4_99,geshi)

#茶之道（云听）
id_4_129 = '16500279462030'
url_4_129 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1724594&columnId=' + id_4_129
Browser.get(url_4_129)
time.sleep(4)

bf_4_129 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_129.text[-1].strip() == '万':
    bofang_4_129 = int(float(bf_4_129.text[:-1])*10000)
else:
    bofang_4_129 = bf_4_129.text

sc_4_129 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_129.text[-1].strip() == '集':
    shangchuan_4_129 = int(float(sc_4_129.text[:-1]))
else:
    shangchuan_4_129 = sc_4_129.text

yuanshimingcheng_4_129 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_129 = yuanshimingcheng_4_129.text
biaotou4_129 = "此次" + name_4_129 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_129)
print('账号为：',name_4_129 + " ；上传量：",shangchuan_4_129 , ' ；播放量：' , bofang_4_129 )

worksheet.write('H85', name_4_129,geshi)
worksheet.write('I85', shangchuan_4_129,geshi)
worksheet.write('J85', bofang_4_129,geshi)

#袁隆平口述自传
id_4_101 = '16217650655220'
url_4_101 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1484011&columnId=' + id_4_101
Browser.get(url_4_101)
time.sleep(4)

bf_4_101 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_101.text[-1].strip() == '万':
    bofang_4_101 = int(float(bf_4_101.text[:-1])*10000)
else:
    bofang_4_101 = bf_4_101.text

sc_4_101 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_101.text[-1].strip() == '集':
    shangchuan_4_101 = int(float(sc_4_101.text[:-1]))
else:
    shangchuan_4_101 = sc_4_101.text

yuanshimingcheng_4_101 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_101 = yuanshimingcheng_4_101.text
biaotou4_101 = "此次" + name_4_101 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_101)
print('账号为：',name_4_101 + " ；上传量：",shangchuan_4_101 , ' ；播放量：' , bofang_4_101 )

worksheet.write('H86', name_4_101,geshi)
worksheet.write('I86', shangchuan_4_101,geshi)
worksheet.write('J86', bofang_4_101,geshi)

#儿童安全情景广播剧【大雄的故事】
id_4_102 = '16218229861660'
url_4_102 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1484200&columnId=' + id_4_102
Browser.get(url_4_102)
time.sleep(4)

bf_4_102 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_102.text[-1].strip() == '万':
    bofang_4_102 = int(float(bf_4_102.text[:-1])*10000)
else:
    bofang_4_102 = bf_4_102.text

sc_4_102 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_102.text[-1].strip() == '集':
    shangchuan_4_102 = int(float(sc_4_102.text[:-1]))
else:
    shangchuan_4_102 = sc_4_102.text

yuanshimingcheng_4_102 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_102 = yuanshimingcheng_4_102.text
biaotou4_102 = "此次" + name_4_102 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_102)
print('账号为：',name_4_102 + " ；上传量：",shangchuan_4_102 , ' ；播放量：' , bofang_4_102 )

worksheet.write('H87', name_4_102,geshi)
worksheet.write('I87', shangchuan_4_102,geshi)
worksheet.write('J87', bofang_4_102,geshi)

#村庄里唱响的红色经典
id_4_103 = '16243407225840'
url_4_103 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1510153&columnId=' + id_4_103
Browser.get(url_4_103)
time.sleep(4)

bf_4_103 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_103.text[-1].strip() == '万':
    bofang_4_103 = int(float(bf_4_103.text[:-1])*10000)
else:
    bofang_4_103 = bf_4_103.text

sc_4_103 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_103.text[-1].strip() == '集':
    shangchuan_4_103 = int(float(sc_4_103.text[:-1]))
else:
    shangchuan_4_103 = sc_4_103.text

yuanshimingcheng_4_103 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_103 = yuanshimingcheng_4_103.text
biaotou4_103 = "此次" + name_4_103 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_103)
print('账号为：',name_4_103 + " ；上传量：",shangchuan_4_103 , ' ；播放量：' , bofang_4_103 )

worksheet.write('H88', name_4_103,geshi)
worksheet.write('I88', shangchuan_4_103,geshi)
worksheet.write('J88', bofang_4_103,geshi)

#康辉聊聊普通话
id_4_104 = '16315441487130'
url_4_104 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1565260&columnId=' + id_4_104
Browser.get(url_4_104)
time.sleep(4)

bf_4_104 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_104.text[-1].strip() == '万':
    bofang_4_104 = int(float(bf_4_104.text[:-1])*10000)
else:
    bofang_4_104 = bf_4_104.text

sc_4_104 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_104.text[-1].strip() == '集':
    shangchuan_4_104 = int(float(sc_4_104.text[:-1]))
else:
    shangchuan_4_104 = sc_4_104.text

yuanshimingcheng_4_104 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_104 = yuanshimingcheng_4_104.text
biaotou4_104 = "此次" + name_4_104 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_104)
print('账号为：',name_4_104 + " ；上传量：",shangchuan_4_104 , ' ；播放量：' , bofang_4_104 )

worksheet.write('H89', name_4_104,geshi)
worksheet.write('I89', shangchuan_4_104,geshi)
worksheet.write('J89', bofang_4_104,geshi)

#了解普通话
id_4_105 = '16318710435100'
url_4_105 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1568277&columnId=' + id_4_105
Browser.get(url_4_105)
time.sleep(4)

bf_4_105 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_105.text[-1].strip() == '万':
    bofang_4_105 = int(float(bf_4_105.text[:-1])*10000)
else:
    bofang_4_105 = bf_4_105.text

sc_4_105 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_105.text[-1].strip() == '集':
    shangchuan_4_105 = int(float(sc_4_105.text[:-1]))
else:
    shangchuan_4_105 = sc_4_105.text

yuanshimingcheng_4_105 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_105 = yuanshimingcheng_4_105.text
biaotou4_105 = "此次" + name_4_105 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_105)
print('账号为：',name_4_105 + " ；上传量：",shangchuan_4_105 , ' ；播放量：' , bofang_4_105 )

worksheet.write('H90', name_4_105,geshi)
worksheet.write('I90', shangchuan_4_105,geshi)
worksheet.write('J90', bofang_4_105,geshi)

#二十四节气 诗词
id_4_106 = '16316957709780'
url_4_106 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1566721&columnId=' + id_4_106
Browser.get(url_4_106)
time.sleep(4)

bf_4_106 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_106.text[-1].strip() == '万':
    bofang_4_106 = int(float(bf_4_106.text[:-1])*10000)
else:
    bofang_4_106 = bf_4_106.text

sc_4_106 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_106.text[-1].strip() == '集':
    shangchuan_4_106 = int(float(sc_4_106.text[:-1]))
else:
    shangchuan_4_106 = sc_4_106.text

yuanshimingcheng_4_106 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_106 = yuanshimingcheng_4_106.text
biaotou4_106 = "此次" + name_4_106 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_106)
print('账号为：',name_4_106 + " ；上传量：",shangchuan_4_106 , ' ；播放量：' , bofang_4_106 )

worksheet.write('H91', name_4_106,geshi)
worksheet.write('I91', shangchuan_4_106,geshi)
worksheet.write('J91', bofang_4_106,geshi)

#二十四节气·诗词（闽方言版）
id_4_107 = '16320668010600'
url_4_107 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1569386&columnId=' + id_4_107
Browser.get(url_4_107)
time.sleep(4)

bf_4_107 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_107.text[-1].strip() == '万':
    bofang_4_107 = int(float(bf_4_107.text[:-1])*10000)
else:
    bofang_4_107 = bf_4_107.text

sc_4_107 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_107.text[-1].strip() == '集':
    shangchuan_4_107 = int(float(sc_4_107.text[:-1]))
else:
    shangchuan_4_107 = sc_4_107.text

yuanshimingcheng_4_107 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_107 = yuanshimingcheng_4_107.text
biaotou4_107 = "此次" + name_4_107 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_107)
print('账号为：',name_4_107 + " ；上传量：",shangchuan_4_107 , ' ；播放量：' , bofang_4_107 )

worksheet.write('H92', name_4_107,geshi)
worksheet.write('I92', shangchuan_4_107,geshi)
worksheet.write('J92', bofang_4_107,geshi)

#二十四节气·诗词（闽客家方言版）
id_4_108 = '16320667951210'
url_4_108 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1569367&columnId=' + id_4_108
Browser.get(url_4_108)
time.sleep(4)

bf_4_108 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_108.text[-1].strip() == '万':
    bofang_4_108 = int(float(bf_4_108.text[:-1])*10000)
else:
    bofang_4_108 = bf_4_108.text

sc_4_108 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_108.text[-1].strip() == '集':
    shangchuan_4_108 = int(float(sc_4_108.text[:-1]))
else:
    shangchuan_4_108 = sc_4_108.text

yuanshimingcheng_4_108 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_108 = yuanshimingcheng_4_108.text
biaotou4_108 = "此次" + name_4_108 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_108)
print('账号为：',name_4_108 + " ；上传量：",shangchuan_4_108 , ' ；播放量：' , bofang_4_108 )

worksheet.write('H93', name_4_108,geshi)
worksheet.write('I93', shangchuan_4_108,geshi)
worksheet.write('J93', bofang_4_108,geshi)

#如果诗人有口音
id_4_109 = '16319595352300'
url_4_109 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1569604&columnId=' + id_4_109
Browser.get(url_4_109)
time.sleep(4)

bf_4_109 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_109.text[-1].strip() == '万':
    bofang_4_109 = int(float(bf_4_109.text[:-1])*10000)
else:
    bofang_4_109 = bf_4_109.text

sc_4_109 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_109.text[-1].strip() == '集':
    shangchuan_4_109 = int(float(sc_4_109.text[:-1]))
else:
    shangchuan_4_109 = sc_4_109.text

yuanshimingcheng_4_109 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_109 = yuanshimingcheng_4_109.text
biaotou4_109 = "此次" + name_4_109 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_109)
print('账号为：',name_4_109 + " ；上传量：",shangchuan_4_109 , ' ；播放量：' , bofang_4_109 )

worksheet.write('H94', name_4_109,geshi)
worksheet.write('I94', shangchuan_4_109,geshi)
worksheet.write('J94', bofang_4_109,geshi)

#听闽方言 赏上古音
id_4_110 = '16326355894810'
url_4_110 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1573271&columnId=' + id_4_110
Browser.get(url_4_110)
time.sleep(4)

bf_4_110 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_110.text[-1].strip() == '万':
    bofang_4_110 = int(float(bf_4_110.text[:-1])*10000)
else:
    bofang_4_110 = bf_4_110.text

sc_4_110 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_110.text[-1].strip() == '集':
    shangchuan_4_110 = int(float(sc_4_110.text[:-1]))
else:
    shangchuan_4_110 = sc_4_110.text

yuanshimingcheng_4_110 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_110 = yuanshimingcheng_4_110.text
biaotou4_110 = "此次" + name_4_110 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_110)
print('账号为：',name_4_110 + " ；上传量：",shangchuan_4_110 , ' ；播放量：' , bofang_4_110 )

worksheet.write('H95', name_4_110,geshi)
worksheet.write('I95', shangchuan_4_110,geshi)
worksheet.write('J95', bofang_4_110,geshi)

#粤说粤嗨森
id_4_111 = '16328049089340'
url_4_111 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1574760&columnId=' + id_4_111
Browser.get(url_4_111)
time.sleep(4)

bf_4_111 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_111.text[-1].strip() == '万':
    bofang_4_111 = int(float(bf_4_111.text[:-1])*10000)
else:
    bofang_4_111 = bf_4_111.text

sc_4_111 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_111.text[-1].strip() == '集':
    shangchuan_4_111 = int(float(sc_4_111.text[:-1]))
else:
    shangchuan_4_111 = sc_4_111.text

yuanshimingcheng_4_111 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_111 = yuanshimingcheng_4_111.text
biaotou4_111 = "此次" + name_4_111 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_111)
print('账号为：',name_4_111 + " ；上传量：",shangchuan_4_111 , ' ；播放量：' , bofang_4_111 )

worksheet.write('H96', name_4_111,geshi)
worksheet.write('I96', shangchuan_4_111,geshi)
worksheet.write('J96', bofang_4_111,geshi)

#科普赣方言
id_4_112 = '16327318954810'
url_4_112 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1574201&columnId=' + id_4_112
Browser.get(url_4_112)
time.sleep(4)

bf_4_112 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_112.text[-1].strip() == '万':
    bofang_4_112 = int(float(bf_4_112.text[:-1])*10000)
else:
    bofang_4_112 = bf_4_112.text

sc_4_112 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_112.text[-1].strip() == '集':
    shangchuan_4_112 = int(float(sc_4_112.text[:-1]))
else:
    shangchuan_4_112 = sc_4_112.text

yuanshimingcheng_4_112 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_112 = yuanshimingcheng_4_112.text
biaotou4_112 = "此次" + name_4_112 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_112)
print('账号为：',name_4_112 + " ；上传量：",shangchuan_4_112 , ' ；播放量：' , bofang_4_112 )

worksheet.write('H97', name_4_112,geshi)
worksheet.write('I97', shangchuan_4_112,geshi)
worksheet.write('J97', bofang_4_112,geshi)

#吴侬软语
id_4_113 = '16327334527220'
url_4_113 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1574242&columnId=' + id_4_113
Browser.get(url_4_113)
time.sleep(4)

bf_4_113 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_113.text[-1].strip() == '万':
    bofang_4_113 = int(float(bf_4_113.text[:-1])*10000)
else:
    bofang_4_113 = bf_4_113.text

sc_4_113 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_113.text[-1].strip() == '集':
    shangchuan_4_113 = int(float(sc_4_113.text[:-1]))
else:
    shangchuan_4_113 = sc_4_113.text

yuanshimingcheng_4_113 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_113 = yuanshimingcheng_4_113.text
biaotou4_113 = "此次" + name_4_113 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_113)
print('账号为：',name_4_113 + " ；上传量：",shangchuan_4_113 , ' ；播放量：' , bofang_4_113 )

worksheet.write('H98', name_4_113,geshi)
worksheet.write('I98', shangchuan_4_113,geshi)
worksheet.write('J98', bofang_4_113,geshi)

#霸蛮湘语
id_4_114 = '16329707516240'
url_4_114 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1576366&columnId=' + id_4_114
Browser.get(url_4_114)
time.sleep(4)

bf_4_114 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_114.text[-1].strip() == '万':
    bofang_4_114 = int(float(bf_4_114.text[:-1])*10000)
else:
    bofang_4_114 = bf_4_114.text

sc_4_114 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_114.text[-1].strip() == '集':
    shangchuan_4_114 = int(float(sc_4_114.text[:-1]))
else:
    shangchuan_4_114 = sc_4_114.text

yuanshimingcheng_4_114 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_114 = yuanshimingcheng_4_114.text
biaotou4_114 = "此次" + name_4_114 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_114)
print('账号为：',name_4_114 + " ；上传量：",shangchuan_4_114 , ' ；播放量：' , bofang_4_114 )

worksheet.write('H99', name_4_114,geshi)
worksheet.write('I99', shangchuan_4_114,geshi)
worksheet.write('J99', bofang_4_114,geshi)

#揭秘客家话
id_4_115 = '16329684222650'
url_4_115 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1576259&columnId=' + id_4_115
Browser.get(url_4_115)
time.sleep(4)

bf_4_115 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_115.text[-1].strip() == '万':
    bofang_4_115 = int(float(bf_4_115.text[:-1])*10000)
else:
    bofang_4_115 = bf_4_115.text

sc_4_115 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_115.text[-1].strip() == '集':
    shangchuan_4_115 = int(float(sc_4_115.text[:-1]))
else:
    shangchuan_4_115 = sc_4_115.text

yuanshimingcheng_4_115 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_115 = yuanshimingcheng_4_115.text
biaotou4_115 = "此次" + name_4_115 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_115)
print('账号为：',name_4_115 + " ；上传量：",shangchuan_4_115 , ' ；播放量：' , bofang_4_115 )

worksheet.write('H100', name_4_115,geshi)
worksheet.write('I100', shangchuan_4_115,geshi)
worksheet.write('J100', bofang_4_115,geshi)

# 水果先生
id_200_43 = '16079494004670'
url_200_43 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1381300&columnId=' + id_200_43
Browser.get(url_200_43)
time.sleep(4)

bf_200_43 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_43.text[-1].strip() == '万':
    bofang_200_43 = int(float(bf_200_43.text[:-1])*10000)
else:
    bofang_200_43 = bf_200_43.text

sc_200_43 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_43.text[-1].strip() == '集':
    shangchuan_200_43 = int(float(sc_200_43.text[:-1]))
else:
    shangchuan_200_43 = sc_200_43.text

yuanshimingcheng_200_43 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_43 = yuanshimingcheng_200_43.text
biaotou200_43 = "此次" + name_200_43 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_43)
print('账号为：',name_200_43 + " ；上传量：",shangchuan_200_43 , ' ；播放量：' , bofang_200_43 )

worksheet.write('H101', name_200_43,geshi)
worksheet.write('I101', shangchuan_200_43,geshi)
worksheet.write('J101', bofang_200_43,geshi)

#咱村的贴心人
id_200_42 = '16093739015550'
url_200_42 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1390225&columnId=' + id_200_42
Browser.get(url_200_42)
time.sleep(4)

bf_200_42 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_42.text[-1].strip() == '万':
    bofang_200_42 = int(float(bf_200_42.text[:-1])*10000)
else:
    bofang_200_42 = bf_200_42.text

sc_200_42 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_42.text[-1].strip() == '集':
    shangchuan_200_42 = int(float(sc_200_42.text[:-1]))
else:
    shangchuan_200_42 = sc_200_42.text

yuanshimingcheng_200_42 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_42 = yuanshimingcheng_200_42.text
biaotou200_42 = "此次" + name_200_42 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_42)
print('账号为：',name_200_42 + " ；上传量：",shangchuan_200_42 , ' ；播放量：' , bofang_200_42 )

worksheet.write('H102', name_200_42,geshi)
worksheet.write('I102', shangchuan_200_42,geshi)
worksheet.write('J102', bofang_200_42,geshi)

#韵味岭南
id_4_118 = '16026685410780'
url_4_118 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1572462&columnId=' + id_4_118
Browser.get(url_4_118)
time.sleep(4)

bf_4_118 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_118.text[-1].strip() == '万':
    bofang_4_118 = int(float(bf_4_118.text[:-1])*10000)
else:
    bofang_4_118 = bf_4_118.text

sc_4_118 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_118.text[-1].strip() == '集':
    shangchuan_4_118 = int(float(sc_4_118.text[:-1]))
else:
    shangchuan_4_118 = sc_4_118.text

yuanshimingcheng_4_118 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_118 = yuanshimingcheng_4_118.text
biaotou4_118 = "此次" + name_4_118 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_118)
print('账号为：',name_4_118 + " ；上传量：",shangchuan_4_118 , ' ；播放量：' , bofang_4_118 )

worksheet.write('H103', name_4_118,geshi)
worksheet.write('I103', shangchuan_4_118,geshi)
worksheet.write('J103', bofang_4_118,geshi)


#魅力中国（合作）
id_4_119 = '15685400767571'
url_4_119 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1125826&columnId=' + id_4_119
Browser.get(url_4_119)
time.sleep(4)

bf_4_119 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_119.text[-1].strip() == '万':
    bofang_4_119 = int(float(bf_4_119.text[:-1])*10000)
else:
    bofang_4_119 = bf_4_119.text

sc_4_119 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_119.text[-1].strip() == '集':
    shangchuan_4_119 = int(float(sc_4_119.text[:-1]))
else:
    shangchuan_4_119 = sc_4_119.text

yuanshimingcheng_4_119 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_119 = yuanshimingcheng_4_119.text
biaotou4_119 = "此次" + name_4_119 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_119)
print('账号为：',name_4_119 + " ；上传量：",shangchuan_4_119 , ' ；播放量：' , bofang_4_119 )

worksheet.write('H104', name_4_119,geshi)
worksheet.write('I104', shangchuan_4_119,geshi)
worksheet.write('J104', bofang_4_119,geshi)

#兜兜白相相
id_4_120 = '15907331585750'
url_4_120 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1326447&columnId=' + id_4_120
Browser.get(url_4_120)
time.sleep(4)

bf_4_120 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_120.text[-1].strip() == '万':
    bofang_4_120 = int(float(bf_4_120.text[:-1])*10000)
else:
    bofang_4_120 = bf_4_120.text

sc_4_120 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_120.text[-1].strip() == '集':
    shangchuan_4_120 = int(float(sc_4_120.text[:-1]))
else:
    shangchuan_4_120 = sc_4_120.text

yuanshimingcheng_4_120 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_120 = yuanshimingcheng_4_120.text
biaotou4_120 = "此次" + name_4_120 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_120)
print('账号为：',name_4_120 + " ；上传量：",shangchuan_4_120 , ' ；播放量：' , bofang_4_120 )

worksheet.write('H105', name_4_120,geshi)
worksheet.write('I105', shangchuan_4_120,geshi)
worksheet.write('J105', bofang_4_120,geshi)

# #听见苏州
id_4_121 = '15855496042000'
url_4_121 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1195643&columnId=' + id_4_121
Browser.get(url_4_121)
time.sleep(4)

bf_4_121 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_121.text[-1].strip() == '万':
    bofang_4_121 = int(float(bf_4_121.text[:-1])*10000)
else:
    bofang_4_121 = bf_4_121.text

sc_4_121 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_121.text[-1].strip() == '集':
    shangchuan_4_121 = int(float(sc_4_121.text[:-1]))
else:
    shangchuan_4_121 = sc_4_121.text

yuanshimingcheng_4_121 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_121 = yuanshimingcheng_4_121.text
biaotou4_121 = "此次" + name_4_121 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_121)
print('账号为：',name_4_121 + " ；上传量：",shangchuan_4_121 , ' ；播放量：' , bofang_4_121 )

worksheet.write('H106', name_4_121,geshi)
worksheet.write('I106', shangchuan_4_121,geshi)
worksheet.write('J106', bofang_4_121,geshi)

#寻城记
id_4_122 = '16286676911530'
url_4_122 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1541712&columnId=' + id_4_122
Browser.get(url_4_122)
time.sleep(4)

bf_4_122 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_122.text[-1].strip() == '万':
    bofang_4_122 = int(float(bf_4_122.text[:-1])*10000)
else:
    bofang_4_122 = bf_4_122.text

sc_4_122 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_122.text[-1].strip() == '集':
    shangchuan_4_122 = int(float(sc_4_122.text[:-1]))
else:
    shangchuan_4_122 = sc_4_122.text

yuanshimingcheng_4_122 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_122 = yuanshimingcheng_4_122.text
biaotou4_122 = "此次" + name_4_122 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_122)
print('账号为：',name_4_122 + " ；上传量：",shangchuan_4_122 , ' ；播放量：' , bofang_4_122 )

worksheet.write('H107', name_4_122,geshi)
worksheet.write('I107', shangchuan_4_122,geshi)
worksheet.write('J107', bofang_4_122,geshi)

#方言与美食
id_4_123 = '16384359213970'
url_4_123 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1618760&columnId=' + id_4_123
Browser.get(url_4_123)
time.sleep(4)

bf_4_123 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_123.text[-1].strip() == '万':
    bofang_4_123 = int(float(bf_4_123.text[:-1])*10000)
else:
    bofang_4_123 = bf_4_123.text

sc_4_123 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_123.text[-1].strip() == '集':
    shangchuan_4_123 = int(float(sc_4_123.text[:-1]))
else:
    shangchuan_4_123 = sc_4_123.text

yuanshimingcheng_4_123 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_123 = yuanshimingcheng_4_123.text
biaotou4_123 = "此次" + name_4_123 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_123)
print('账号为：',name_4_123 + " ；上传量：",shangchuan_4_123 , ' ；播放量：' , bofang_4_123 )

worksheet.write('H108', name_4_123,geshi)
worksheet.write('I108', shangchuan_4_123,geshi)
worksheet.write('J108', bofang_4_123,geshi)

#李大勺开小灶
id_4_124 = '16399874429520'
url_4_124 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1630343&columnId=' + id_4_124
Browser.get(url_4_124)
time.sleep(4)

bf_4_124 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_124.text[-1].strip() == '万':
    bofang_4_124 = int(float(bf_4_124.text[:-1])*10000)
else:
    bofang_4_124 = bf_4_124.text

sc_4_124 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_124.text[-1].strip() == '集':
    shangchuan_4_124 = int(float(sc_4_124.text[:-1]))
else:
    shangchuan_4_124 = sc_4_124.text

yuanshimingcheng_4_124 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_124 = yuanshimingcheng_4_124.text
biaotou4_124 = "此次" + name_4_124 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_124)
print('账号为：',name_4_124 + " ；上传量：",shangchuan_4_124 , ' ；播放量：' , bofang_4_124 )

worksheet.write('H109', name_4_124,geshi)
worksheet.write('I109', shangchuan_4_124,geshi)
worksheet.write('J109', bofang_4_124,geshi)

#中华老字号
id_4_125 = '15682080945972'
url_4_125 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=531073&columnId=' + id_4_125
Browser.get(url_4_125)
time.sleep(4)

bf_4_125 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_125.text[-1].strip() == '万':
    bofang_4_125 = int(float(bf_4_125.text[:-1])*10000)
else:
    bofang_4_125 = bf_4_125.text

sc_4_125 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_125.text[-1].strip() == '集':
    shangchuan_4_125 = int(float(sc_4_125.text[:-1]))
else:
    shangchuan_4_125 = sc_4_125.text

yuanshimingcheng_4_125 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_125 = yuanshimingcheng_4_125.text
biaotou4_125 = "此次" + name_4_125 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_125)
print('账号为：',name_4_125 + " ；上传量：",shangchuan_4_125 , ' ；播放量：' , bofang_4_125 )

worksheet.write('H110', name_4_125,geshi)
worksheet.write('I110', shangchuan_4_125,geshi)
worksheet.write('J110', bofang_4_125,geshi)

#牦牛妈妈
id_4_126 = '16405840897910'
url_4_126 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1637630&columnId=' + id_4_126
Browser.get(url_4_126)
time.sleep(4)

bf_4_126 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_126.text[-1].strip() == '万':
    bofang_4_126 = int(float(bf_4_126.text[:-1])*10000)
else:
    bofang_4_126 = bf_4_126.text

sc_4_126 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_126.text[-1].strip() == '集':
    shangchuan_4_126 = int(float(sc_4_126.text[:-1]))
else:
    shangchuan_4_126 = sc_4_126.text

yuanshimingcheng_4_126 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_126 = yuanshimingcheng_4_126.text
biaotou4_126 = "此次" + name_4_126 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_126)
print('账号为：',name_4_126 + " ；上传量：",shangchuan_4_126 , ' ；播放量：' , bofang_4_126 )

worksheet.write('H111', name_4_126,geshi)
worksheet.write('I111', shangchuan_4_126,geshi)
worksheet.write('J111', bofang_4_126,geshi)

#一门三院士 共育稻花香
id_4_127 = '16400623874320'
url_4_127 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1637681&columnId=' + id_4_127
Browser.get(url_4_127)
time.sleep(4)

bf_4_127 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_127.text[-1].strip() == '万':
    bofang_4_127 = int(float(bf_4_127.text[:-1])*10000)
else:
    bofang_4_127 = bf_4_127.text

sc_4_127 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_127.text[-1].strip() == '集':
    shangchuan_4_127 = int(float(sc_4_127.text[:-1]))
else:
    shangchuan_4_127 = sc_4_127.text

yuanshimingcheng_4_127 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_127 = yuanshimingcheng_4_127.text
biaotou4_127 = "此次" + name_4_127 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_127)
print('账号为：',name_4_127 + " ；上传量：",shangchuan_4_127 , ' ；播放量：' , bofang_4_127 )

worksheet.write('H112', name_4_127,geshi)
worksheet.write('I112', shangchuan_4_127,geshi)
worksheet.write('J112', bofang_4_127,geshi)

#大国农数
id_4_128 = '16405841073700'
url_4_128 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1642705&columnId=' + id_4_128
Browser.get(url_4_128)
time.sleep(4)

bf_4_128 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_128.text[-1].strip() == '万':
    bofang_4_128 = int(float(bf_4_128.text[:-1])*10000)
else:
    bofang_4_128 = bf_4_128.text

sc_4_128 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_128.text[-1].strip() == '集':
    shangchuan_4_128 = int(float(sc_4_128.text[:-1]))
else:
    shangchuan_4_128 = sc_4_128.text

yuanshimingcheng_4_128 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_128 = yuanshimingcheng_4_128.text + '(云听)'
biaotou4_128 = "此次" + name_4_128 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_128)
print('账号为：',name_4_128 + " ；上传量：",shangchuan_4_128 , ' ；播放量：' , bofang_4_128 )

worksheet.write('H113', name_4_128,geshi)
worksheet.write('I113', shangchuan_4_128,geshi)
worksheet.write('J113', bofang_4_128,geshi)

#丰晚bgm（云听）
id_4_132 = '16323765648820'
url_4_132 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1571435&columnId=' + id_4_132
Browser.get(url_4_132)
time.sleep(4)

bf_4_132 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_132.text[-1].strip() == '万':
    bofang_4_132 = int(float(bf_4_132.text[:-1])*10000)
else:
    bofang_4_132 = bf_4_132.text

sc_4_132 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_132.text[-1].strip() == '集':
    shangchuan_4_132 = int(float(sc_4_132.text[:-1]))
else:
    shangchuan_4_132 = sc_4_132.text

yuanshimingcheng_4_132 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_132 = yuanshimingcheng_4_132.text
biaotou4_132 = "此次" + name_4_132 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_132)
print('账号为：',name_4_132 + " ；上传量：",shangchuan_4_132 , ' ；播放量：' , bofang_4_132 )

worksheet.write('H114', name_4_132,geshi)
worksheet.write('I114', shangchuan_4_132,geshi)
worksheet.write('J114', bofang_4_132,geshi)

#田野美味（云听）
id_4_133 = '16475853813290'
url_4_133 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1687843&columnId=' + id_4_133
Browser.get(url_4_133)
time.sleep(4)

bf_4_133 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_133.text[-1].strip() == '万':
    bofang_4_133 = int(float(bf_4_133.text[:-1])*10000)
else:
    bofang_4_133 = bf_4_133.text

sc_4_133 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_133.text[-1].strip() == '集':
    shangchuan_4_133 = int(float(sc_4_133.text[:-1]))
else:
    shangchuan_4_133 = sc_4_133.text

yuanshimingcheng_4_133 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_133 = yuanshimingcheng_4_133.text
biaotou4_133 = "此次" + name_4_133 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_133)
print('账号为：',name_4_133 + " ；上传量：",shangchuan_4_133 , ' ；播放量：' , bofang_4_133 )

worksheet.write('H115', name_4_133,geshi)
worksheet.write('I115', shangchuan_4_133,geshi)
worksheet.write('J115', bofang_4_133,geshi)


#三农绿厅（云听）
id_4_134 = '16464009101010'
url_4_134 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1683275&columnId=' + id_4_134
Browser.get(url_4_134)
time.sleep(4)

bf_4_134 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_134.text[-1].strip() == '万':
    bofang_4_134 = int(float(bf_4_134.text[:-1])*10000)
else:
    bofang_4_134 = bf_4_134.text

sc_4_134 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_134.text[-1].strip() == '集':
    shangchuan_4_134 = int(float(sc_4_134.text[:-1]))
else:
    shangchuan_4_134 = sc_4_134.text

yuanshimingcheng_4_134 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_134 = yuanshimingcheng_4_134.text
biaotou4_134 = "此次" + name_4_134 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_134)
print('账号为：',name_4_134 + " ；上传量：",shangchuan_4_134 , ' ；播放量：' , bofang_4_134 )

worksheet.write('H116', name_4_134,geshi)
worksheet.write('I116', shangchuan_4_134,geshi)
worksheet.write('J116', bofang_4_134,geshi)

#本草华夏（云听）
id_4_135 = '16404832107760'
url_4_135 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1675539&columnId=' + id_4_135
Browser.get(url_4_135)
time.sleep(4)

bf_4_135 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_135.text[-1].strip() == '万':
    bofang_4_135 = int(float(bf_4_135.text[:-1])*10000)
else:
    bofang_4_135 = bf_4_135.text

sc_4_135 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_135.text[-1].strip() == '集':
    shangchuan_4_135 = int(float(sc_4_135.text[:-1]))
else:
    shangchuan_4_135 = sc_4_135.text

yuanshimingcheng_4_135 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_135 = yuanshimingcheng_4_135.text
biaotou4_135 = "此次" + name_4_135 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_135)
print('账号为：',name_4_135 + " ；上传量：",shangchuan_4_135 , ' ；播放量：' , bofang_4_135 )

worksheet.write('H117', name_4_135,geshi)
worksheet.write('I117', shangchuan_4_135,geshi)
worksheet.write('J117', bofang_4_135,geshi)

#乡音里的中国节（云听）
id_4_131 = '16433516506510'
url_4_131 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1656740&columnId=' + id_4_131
Browser.get(url_4_131)
time.sleep(4)

bf_4_131 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_131.text[-1].strip() == '万':
    bofang_4_131 = int(float(bf_4_131.text[:-1])*10000)
else:
    bofang_4_131 = bf_4_131.text

sc_4_131 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_131.text[-1].strip() == '集':
    shangchuan_4_131 = int(float(sc_4_131.text[:-1]))
else:
    shangchuan_4_131 = sc_4_131.text

yuanshimingcheng_4_131 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_131 = yuanshimingcheng_4_131.text
biaotou4_131 = "此次" + name_4_131 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_131)
print('账号为：',name_4_131 + " ；上传量：",shangchuan_4_131 , ' ；播放量：' , bofang_4_131 )

worksheet.write('H118', name_4_131,geshi)
worksheet.write('I118', shangchuan_4_131,geshi)
worksheet.write('J118', bofang_4_131,geshi)


#慧思带你聊心理（云听）
id_4_136 = '16637507045570'
url_4_136 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1946733&columnId=' + id_4_136
Browser.get(url_4_136)
time.sleep(4)

bf_4_136 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_136.text[-1].strip() == '万':
    bofang_4_136 = int(float(bf_4_136.text[:-1])*10000)
else:
    bofang_4_136 = bf_4_136.text

sc_4_136 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_136.text[-1].strip() == '集':
    shangchuan_4_136 = int(float(sc_4_136.text[:-1]))
else:
    shangchuan_4_136 = sc_4_136.text

yuanshimingcheng_4_136 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_136 = yuanshimingcheng_4_136.text
biaotou4_131 = "此次" + name_4_136 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_131)
print('账号为：',name_4_136 + " ；上传量：",shangchuan_4_136 , ' ；播放量：' , bofang_4_136 )

worksheet.write('H119', name_4_136,geshi)
worksheet.write('I119', shangchuan_4_136,geshi)
worksheet.write('J119', bofang_4_136,geshi)

# -------------------------------------------------------------------------------------------------------------------------
#茶有道 喝茶有门道
id_200_1 = '17085043205840'
url_200_1 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2370322&columnId=' + id_200_1
Browser.get(url_200_1)
time.sleep(4)

bf_200_1 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_1.text[-1].strip() == '万':
    bofang_200_1 = int(float(bf_200_1.text[:-1])*10000)
else:
    bofang_200_1 = bf_200_1.text

sc_200_1 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_1.text[-1].strip() == '集':
    shangchuan_200_1 = int(float(sc_200_1.text[:-1]))
else:
    shangchuan_200_1 = sc_200_1.text

yuanshimingcheng_200_1 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_1 = yuanshimingcheng_200_1.text
biaotou200_1 = "此次" + name_200_1 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_1)
print('账号为：',name_200_1 + " ；上传量：",shangchuan_200_1 , ' ；播放量：' , bofang_200_1 )

worksheet.write('H120', name_200_1,geshi)
worksheet.write('I120', shangchuan_200_1,geshi)
worksheet.write('J120', bofang_200_1,geshi)

#我是探险家
id_200_2 = '17113509883580'
url_200_2 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2368004&columnId=' + id_200_2
Browser.get(url_200_2)
time.sleep(4)

bf_200_2 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_2.text[-1].strip() == '万':
    bofang_200_2 = int(float(bf_200_2.text[:-1])*10000)
else:
    bofang_200_2 = bf_200_2.text

sc_200_2 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_2.text[-1].strip() == '集':
    shangchuan_200_2 = int(float(sc_200_2.text[:-1]))
else:
    shangchuan_200_2 = sc_200_2.text

yuanshimingcheng_200_2 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_2 = yuanshimingcheng_200_2.text
biaotou200_2 = "此次" + name_200_2 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_2)
print('账号为：',name_200_2 + " ；上传量：",shangchuan_200_2 , ' ；播放量：' , bofang_200_2 )

worksheet.write('H121', name_200_2,geshi)
worksheet.write('I121', shangchuan_200_2,geshi)
worksheet.write('J121', bofang_200_2,geshi)

#古农书带你去穿越
id_200_3 = '17127307543550'
url_200_3 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2372642&columnId=' + id_200_3
Browser.get(url_200_3)
time.sleep(4)

bf_200_3 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_3.text[-1].strip() == '万':
    bofang_200_3 = int(float(bf_200_3.text[:-1])*10000)
else:
    bofang_200_3 = bf_200_3.text

sc_200_3 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_3.text[-1].strip() == '集':
    shangchuan_200_3 = int(float(sc_200_3.text[:-1]))
else:
    shangchuan_200_3 = sc_200_3.text

yuanshimingcheng_200_3 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_3 = yuanshimingcheng_200_3.text
biaotou200_3 = "此次" + name_200_3 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_3)
print('账号为：',name_200_3 + " ；上传量：",shangchuan_200_3 , ' ；播放量：' , bofang_200_3 )

worksheet.write('H122', name_200_3,geshi)
worksheet.write('I122', shangchuan_200_3,geshi)
worksheet.write('J122', bofang_200_3,geshi)

#越听越饿的古代美食
id_200_4 = '17121273977060'
url_200_4 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2369201&columnId=' + id_200_4
Browser.get(url_200_4)
time.sleep(4)

bf_200_4 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_4.text[-1].strip() == '万':
    bofang_200_4 = int(float(bf_200_4.text[:-1])*10000)
else:
    bofang_200_4 = bf_200_4.text

sc_200_4 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_4.text[-1].strip() == '集':
    shangchuan_200_4 = int(float(sc_200_4.text[:-1]))
else:
    shangchuan_200_4 = sc_200_4.text

yuanshimingcheng_200_4 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_4 = yuanshimingcheng_200_4.text
biaotou200_4 = "此次" + name_200_4 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_4)
print('账号为：',name_200_4 + " ；上传量：",shangchuan_200_4 , ' ；播放量：' , bofang_200_4 )

worksheet.write('H123', name_200_4,geshi)
worksheet.write('I123', shangchuan_200_4,geshi)
worksheet.write('J123', bofang_200_4,geshi)

#谈谷论今——谷物塑造的人类史
id_200_5 = '17121274498340'
url_200_5 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2375543&columnId=' + id_200_5
Browser.get(url_200_5)
time.sleep(4)

bf_200_5 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_5.text[-1].strip() == '万':
    bofang_200_5 = int(float(bf_200_5.text[:-1])*10000)
else:
    bofang_200_5 = bf_200_5.text

sc_200_5 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_5.text[-1].strip() == '集':
    shangchuan_200_5 = int(float(sc_200_5.text[:-1]))
else:
    shangchuan_200_5 = sc_200_5.text

yuanshimingcheng_200_5 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_5 = yuanshimingcheng_200_5.text
biaotou200_5 = "此次" + name_200_5 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_5)
print('账号为：',name_200_5 + " ；上传量：",shangchuan_200_5 , ' ；播放量：' , bofang_200_5 )

worksheet.write('H124', name_200_5,geshi)
worksheet.write('I124', shangchuan_200_5,geshi)
worksheet.write('J124', bofang_200_5,geshi)

#1000个婚姻故事中的AB面
id_200_6 = '17153919362900'
url_200_6 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2396900&columnId=' + id_200_6
Browser.get(url_200_6)
time.sleep(4)

bf_200_6 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_6.text[-1].strip() == '万':
    bofang_200_6 = int(float(bf_200_6.text[:-1])*10000)
else:
    bofang_200_6 = bf_200_6.text

sc_200_6 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_6.text[-1].strip() == '集':
    shangchuan_200_6 = int(float(sc_200_6.text[:-1]))
else:
    shangchuan_200_6 = sc_200_6.text

yuanshimingcheng_200_6 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_6 = yuanshimingcheng_200_6.text
biaotou200_6 = "此次" + name_200_6 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_6)
print('账号为：',name_200_6 + " ；上传量：",shangchuan_200_6 , ' ；播放量：' , bofang_200_6 )

worksheet.write('H125', name_200_6,geshi)
worksheet.write('I125', shangchuan_200_6,geshi)
worksheet.write('J125', bofang_200_6,geshi)

#越吃越有范儿
id_200_7 = '17173837891430'
url_200_7 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2410236&columnId=' + id_200_7
Browser.get(url_200_7)
time.sleep(4)

bf_200_7 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_7.text[-1].strip() == '万':
    bofang_200_7 = int(float(bf_200_7.text[:-1])*10000)
else:
    bofang_200_7 = bf_200_7.text

sc_200_7 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_7.text[-1].strip() == '集':
    shangchuan_200_7 = int(float(sc_200_7.text[:-1]))
else:
    shangchuan_200_7 = sc_200_7.text

yuanshimingcheng_200_7 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_7 = yuanshimingcheng_200_7.text
biaotou200_7 = "此次" + name_200_7 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_7)
print('账号为：',name_200_7 + " ；上传量：",shangchuan_200_7 , ' ；播放量：' , bofang_200_7 )

worksheet.write('H126', name_200_7,geshi)
worksheet.write('I126', shangchuan_200_7,geshi)
worksheet.write('J126', bofang_200_7,geshi)

#探秘千年古树
id_200_8 = '17178993813150'
url_200_8 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2419406&columnId=' + id_200_8
Browser.get(url_200_8)
time.sleep(4)

bf_200_8 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_8.text[-1].strip() == '万':
    bofang_200_8 = int(float(bf_200_8.text[:-1])*10000)
else:
    bofang_200_8 = bf_200_8.text

sc_200_8 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_8.text[-1].strip() == '集':
    shangchuan_200_8 = int(float(sc_200_8.text[:-1]))
else:
    shangchuan_200_8 = sc_200_8.text

yuanshimingcheng_200_8 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_8 = yuanshimingcheng_200_8.text
biaotou200_8 = "此次" + name_200_8 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_8)
print('账号为：',name_200_8 + " ；上传量：",shangchuan_200_8 , ' ；播放量：' , bofang_200_8 )

worksheet.write('H127', name_200_8,geshi)
worksheet.write('I127', shangchuan_200_8,geshi)
worksheet.write('J127', bofang_200_8,geshi)

#只有香如故
id_200_9 = '17204250706630'
url_200_9 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2431761&columnId=' + id_200_9
Browser.get(url_200_9)
time.sleep(4)

bf_200_9 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_9.text[-1].strip() == '万':
    bofang_200_9 = int(float(bf_200_9.text[:-1])*10000)
else:
    bofang_200_9 = bf_200_9.text

sc_200_9 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_9.text[-1].strip() == '集':
    shangchuan_200_9 = int(float(sc_200_9.text[:-1]))
else:
    shangchuan_200_9 = sc_200_9.text

yuanshimingcheng_200_9 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_9 = yuanshimingcheng_200_9.text
biaotou200_9 = "此次" + name_200_9 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_9)
print('账号为：',name_200_9 + " ；上传量：",shangchuan_200_9 , ' ；播放量：' , bofang_200_9 )

worksheet.write('H128', name_200_9,geshi)
worksheet.write('I128', shangchuan_200_9,geshi)
worksheet.write('J128', bofang_200_9,geshi)

# #听见非遗
# id_200_10 = '17224056000060'
# url_200_10 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2446361&columnId=' + id_200_10
# Browser.get(url_200_10)
# time.sleep(4)
#
# bf_200_10 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
# if bf_200_10.text[-1].strip() == '万':
#     bofang_200_10 = int(float(bf_200_10.text[:-1])*10000)
# else:
#     bofang_200_10 = bf_200_10.text
#
# sc_200_10 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
# if sc_200_10.text[-1].strip() == '集':
#     shangchuan_200_10 = int(float(sc_200_10.text[:-1]))
# else:
#     shangchuan_200_10 = sc_200_10.text
#
# yuanshimingcheng_200_10 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
# name_200_10 = yuanshimingcheng_200_10.text
# biaotou200_10 = "此次" + name_200_10 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))
#
# print (biaotou200_10)
# print('账号为：',name_200_10 + " ；上传量：",shangchuan_200_10 , ' ；播放量：' , bofang_200_10 )
#
# worksheet.write('H129', name_200_10,geshi)
# worksheet.write('I129', shangchuan_200_10,geshi)
# worksheet.write('J129', bofang_200_10,geshi)

#听见乡村
#健康到家(重播)(回听)
id_200_15 = '17198965407430'
url_200_15 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2442074&columnId=' + id_200_15
Browser.get(url_200_15)
time.sleep(4)

bf_200_15 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_15.text[-1].strip() == '万':
    bofang_200_15 = int(float(bf_200_15.text[:-1])*10000)
else:
    bofang_200_15 = bf_200_15.text

sc_200_15 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_15.text[-1].strip() == '集':
    shangchuan_200_15 = int(float(sc_200_15.text[:-1]))
else:
    shangchuan_200_15 = sc_200_15.text

yuanshimingcheng_200_15 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_15 = yuanshimingcheng_200_15.text + '(回听)'
biaotou200_15 = "此次" + name_200_15 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_15)
print('账号为：',name_200_15 + " ；上传量：",shangchuan_200_15 , ' ；播放量：' , bofang_200_15 )

worksheet.write('H130', name_200_15,geshi)
worksheet.write('I130', shangchuan_200_15,geshi)
worksheet.write('J130', bofang_200_15,geshi)

#健康开聊吧
id_200_16 = '17243930974440'
url_200_16 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2463091&columnId=' + id_200_16
Browser.get(url_200_16)
time.sleep(4)

bf_200_16 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_16.text[-1].strip() == '万':
    bofang_200_16 = int(float(bf_200_16.text[:-1])*10000)
else:
    bofang_200_16 = bf_200_16.text

sc_200_16 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_16.text[-1].strip() == '集':
    shangchuan_200_16 = int(float(sc_200_16.text[:-1]))
else:
    shangchuan_200_16 = sc_200_16.text

yuanshimingcheng_200_16 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_16 = yuanshimingcheng_200_16.text
biaotou200_16 = "此次" + name_200_16 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_16)
print('账号为：',name_200_16 + " ；上传量：",shangchuan_200_16 , ' ；播放量：' , bofang_200_16 )

worksheet.write('H131', name_200_16,geshi)
worksheet.write('I131', shangchuan_200_16,geshi)
worksheet.write('J131', bofang_200_16,geshi)

#农情档案馆
id_200_17 = '17259728585820'
url_200_17 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2474961&columnId=' + id_200_17
Browser.get(url_200_17)
time.sleep(4)

bf_200_17 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_17.text[-1].strip() == '万':
    bofang_200_17 = int(float(bf_200_17.text[:-1])*10000)
else:
    bofang_200_17 = bf_200_17.text

sc_200_17 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_17.text[-1].strip() == '集':
    shangchuan_200_17 = int(float(sc_200_17.text[:-1]))
else:
    shangchuan_200_17 = sc_200_17.text

yuanshimingcheng_200_17 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_17 = yuanshimingcheng_200_17.text
biaotou200_17 = "此次" + name_200_17 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_17)
print('账号为：',name_200_17 + " ；上传量：",shangchuan_200_17 , ' ；播放量：' , bofang_200_17 )

worksheet.write('H132', name_200_17,geshi)
worksheet.write('I132', shangchuan_200_17,geshi)
worksheet.write('J132', bofang_200_17,geshi)

#文学与人生
id_200_18 = '17260563602160'
url_200_18 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2480061&columnId=' + id_200_18
Browser.get(url_200_18)
time.sleep(4)

bf_200_18 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_18.text[-1].strip() == '万':
    bofang_200_18 = int(float(bf_200_18.text[:-1])*10000)
else:
    bofang_200_18 = bf_200_18.text

sc_200_18 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_18.text[-1].strip() == '集':
    shangchuan_200_18 = int(float(sc_200_18.text[:-1]))
else:
    shangchuan_200_18 = sc_200_18.text

yuanshimingcheng_200_18 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_18 = yuanshimingcheng_200_18.text
biaotou200_18 = "此次" + name_200_18 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_18)
print('账号为：',name_200_18 + " ；上传量：",shangchuan_200_18 , ' ；播放量：' , bofang_200_18 )

worksheet.write('H133', name_200_18,geshi)
worksheet.write('I133', shangchuan_200_18,geshi)
worksheet.write('J133', bofang_200_18,geshi)

#农忙农闲
id_200_19 = '17265591454650'
url_200_19 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2491016&columnId=' + id_200_19
Browser.get(url_200_19)
time.sleep(4)

bf_200_19 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_19.text[-1].strip() == '万':
    bofang_200_19 = int(float(bf_200_19.text[:-1])*10000)
else:
    bofang_200_19 = bf_200_19.text

sc_200_19 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_19.text[-1].strip() == '集':
    shangchuan_200_19 = int(float(sc_200_19.text[:-1]))
else:
    shangchuan_200_19 = sc_200_19.text

yuanshimingcheng_200_19 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_19 = yuanshimingcheng_200_19.text
biaotou200_19 = "此次" + name_200_19 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_19)
print('账号为：',name_200_19 + " ；上传量：",shangchuan_200_19 , ' ；播放量：' , bofang_200_19 )

worksheet.write('H134', name_200_19,geshi)
worksheet.write('I134', shangchuan_200_19,geshi)
worksheet.write('J134', bofang_200_19,geshi)

#我和大自然有个约会
id_200_20 = '17288872370880'
url_200_20 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2499355&columnId=' + id_200_20
Browser.get(url_200_20)
time.sleep(4)

bf_200_20 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_20.text[-1].strip() == '万':
    bofang_200_20 = int(float(bf_200_20.text[:-1])*10000)
else:
    bofang_200_20 = bf_200_20.text

sc_200_20 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_20.text[-1].strip() == '集':
    shangchuan_200_20 = int(float(sc_200_20.text[:-1]))
else:
    shangchuan_200_20 = sc_200_20.text

yuanshimingcheng_200_20 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_20 = yuanshimingcheng_200_20.text
biaotou200_20 = "此次" + name_200_20 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_20)
print('账号为：',name_200_20 + " ；上传量：",shangchuan_200_20 , ' ；播放量：' , bofang_200_20 )

worksheet.write('H135', name_200_20,geshi)
worksheet.write('I135', shangchuan_200_20,geshi)
worksheet.write('J135', bofang_200_20,geshi)


#牙医私房课
id_200_21 = '17271681489670'
url_200_21 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2491000&columnId=' + id_200_21
Browser.get(url_200_21)
time.sleep(4)

bf_200_21 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_21.text[-1].strip() == '万':
    bofang_200_21 = int(float(bf_200_21.text[:-1])*10000)
else:
    bofang_200_21 = bf_200_21.text

sc_200_21 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_21.text[-1].strip() == '集':
    shangchuan_200_21 = int(float(sc_200_21.text[:-1]))
else:
    shangchuan_200_21 = sc_200_21.text

yuanshimingcheng_200_21 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_21 = yuanshimingcheng_200_21.text
biaotou200_21 = "此次" + name_200_21 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_21)
print('账号为：',name_200_21 + " ；上传量：",shangchuan_200_21 , ' ；播放量：' , bofang_200_21 )

worksheet.write('H136', name_200_21,geshi)
worksheet.write('I136', shangchuan_200_21,geshi)
worksheet.write('J136', bofang_200_21,geshi)



#花开的声音
id_200_22 = '17305624929640'
url_200_22 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2519565&columnId=' + id_200_22
Browser.get(url_200_22)
time.sleep(4)

bf_200_22 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_22.text[-1].strip() == '万':
    bofang_200_22 = int(float(bf_200_22.text[:-1])*10000)
else:
    bofang_200_22 = bf_200_22.text

sc_200_22 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_22.text[-1].strip() == '集':
    shangchuan_200_22 = int(float(sc_200_22.text[:-1]))
else:
    shangchuan_200_22 = sc_200_22.text

yuanshimingcheng_200_22 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_22 = yuanshimingcheng_200_22.text
biaotou200_22 = "此次" + name_200_22 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_22)
print('账号为：',name_200_22 + " ；上传量：",shangchuan_200_22 , ' ；播放量：' , bofang_200_22 )

worksheet.write('H137', name_200_22,geshi)
worksheet.write('I137', shangchuan_200_22,geshi)
worksheet.write('J137', bofang_200_22,geshi)




#动物研究所
id_200_23 = '17314679642310'
url_200_23 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2526970&columnId=' + id_200_23
Browser.get(url_200_23)
time.sleep(4)

bf_200_23 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_23.text[-1].strip() == '万':
    bofang_200_23 = int(float(bf_200_23.text[:-1])*10000)
else:
    bofang_200_23 = bf_200_23.text

sc_200_23 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_23.text[-1].strip() == '集':
    shangchuan_200_23 = int(float(sc_200_23.text[:-1]))
else:
    shangchuan_200_23 = sc_200_23.text

yuanshimingcheng_200_23 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_23 = yuanshimingcheng_200_23.text
biaotou200_23 = "此次" + name_200_23 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_23)
print('账号为：',name_200_23 + " ；上传量：",shangchuan_200_23 , ' ；播放量：' , bofang_200_23 )

worksheet.write('H138', name_200_23,geshi)
worksheet.write('I138', shangchuan_200_23,geshi)
worksheet.write('J138', bofang_200_23,geshi)


#幸福公开课
id_200_24 = '17319844973330'
url_200_24 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2534483&columnId=' + id_200_24
Browser.get(url_200_24)
time.sleep(4)

bf_200_24 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_24.text[-1].strip() == '万':
    bofang_200_24 = int(float(bf_200_24.text[:-1])*10000)
else:
    bofang_200_24 = bf_200_24.text

sc_200_24 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_24.text[-1].strip() == '集':
    shangchuan_200_24 = int(float(sc_200_24.text[:-1]))
else:
    shangchuan_200_24 = sc_200_24.text

yuanshimingcheng_200_24 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_24 = yuanshimingcheng_200_24.text
biaotou200_24 = "此次" + name_200_24 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_24)
print('账号为：',name_200_24 + " ；上传量：",shangchuan_200_24 , ' ；播放量：' , bofang_200_24 )

worksheet.write('H139', name_200_24,geshi)
worksheet.write('I139', shangchuan_200_24,geshi)
worksheet.write('J139', bofang_200_24,geshi)

#神奇动物在这里
id_200_25 = '17359757609470'
url_200_25 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2566566&columnId=' + id_200_25
Browser.get(url_200_25)
time.sleep(4)

bf_200_25 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_25.text[-1].strip() == '万':
    bofang_200_25 = int(float(bf_200_25.text[:-1])*10000)
else:
    bofang_200_25 = bf_200_25.text

sc_200_25 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_25.text[-1].strip() == '集':
    shangchuan_200_25 = int(float(sc_200_25.text[:-1]))
else:
    shangchuan_200_25 = sc_200_25.text

yuanshimingcheng_200_25 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_25 = yuanshimingcheng_200_25.text
biaotou200_25 = "此次" + name_200_25 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_25)
print('账号为：',name_200_25 + " ；上传量：",shangchuan_200_25 , ' ；播放量：' , bofang_200_25 )

worksheet.write('H140', name_200_25,geshi)
worksheet.write('I140', shangchuan_200_25,geshi)
worksheet.write('J140', bofang_200_25,geshi)



#熟语里的黄河
id_200_26 = '17363451862810'
url_200_26 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2569194&columnId=' + id_200_26
Browser.get(url_200_26)
time.sleep(4)

bf_200_26 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_26.text[-1].strip() == '万':
    bofang_200_26 = int(float(bf_200_26.text[:-1])*10000)
else:
    bofang_200_26 = bf_200_26.text

sc_200_26 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_26.text[-1].strip() == '集':
    shangchuan_200_26 = int(float(sc_200_26.text[:-1]))
else:
    shangchuan_200_26 = sc_200_26.text

yuanshimingcheng_200_26 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_26 = yuanshimingcheng_200_26.text
biaotou200_26 = "此次" + name_200_26 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_26)
print('账号为：',name_200_26 + " ；上传量：",shangchuan_200_26 , ' ；播放量：' , bofang_200_26 )

worksheet.write('H141', name_200_26,geshi)
worksheet.write('I141', shangchuan_200_26,geshi)
worksheet.write('J141', bofang_200_26,geshi)


#村村向上
id_200_27 = '17369963624080'
url_200_27 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2574600&columnId=' + id_200_27
Browser.get(url_200_27)
time.sleep(4)

bf_200_27 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_27.text[-1].strip() == '万':
    bofang_200_27 = int(float(bf_200_27.text[:-1])*10000)
else:
    bofang_200_27 = bf_200_27.text

sc_200_27 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_27.text[-1].strip() == '集':
    shangchuan_200_27 = int(float(sc_200_27.text[:-1]))
else:
    shangchuan_200_27 = sc_200_27.text

yuanshimingcheng_200_27 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_27 = yuanshimingcheng_200_27.text
biaotou200_27 = "此次" + name_200_27 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_27)
print('账号为：',name_200_27 + " ；上传量：",shangchuan_200_27 , ' ；播放量：' , bofang_200_27 )

worksheet.write('H142', name_200_27,geshi)
worksheet.write('I142', shangchuan_200_27,geshi)
worksheet.write('J142', bofang_200_27,geshi)


#遇见非遗
id_200_28 = '17319964477000'
url_200_28 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2584500&columnId=' + id_200_28
Browser.get(url_200_28)
time.sleep(4)

bf_200_28 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_28.text[-1].strip() == '万':
    bofang_200_28 = int(float(bf_200_28.text[:-1])*10000)
else:
    bofang_200_28 = bf_200_28.text

sc_200_28 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_28.text[-1].strip() == '集':
    shangchuan_200_28 = int(float(sc_200_28.text[:-1]))
else:
    shangchuan_200_28 = sc_200_28.text

yuanshimingcheng_200_28 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_28 = yuanshimingcheng_200_28.text
biaotou200_28 = "此次" + name_200_28 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_28)
print('账号为：',name_200_28 + " ；上传量：",shangchuan_200_28 , ' ；播放量：' , bofang_200_28 )

worksheet.write('H143', name_200_28,geshi)
worksheet.write('I143', shangchuan_200_28,geshi)
worksheet.write('J143', bofang_200_28,geshi)


#围炉话三农
id_200_29 = '17370962826970'
url_200_29 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2603496&columnId=' + id_200_29
Browser.get(url_200_29)
time.sleep(4)

bf_200_29 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_29.text[-1].strip() == '万':
    bofang_200_29 = int(float(bf_200_29.text[:-1])*10000)
else:
    bofang_200_29 = bf_200_29.text

sc_200_29 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_29.text[-1].strip() == '集':
    shangchuan_200_29 = int(float(sc_200_29.text[:-1]))
else:
    shangchuan_200_29 = sc_200_29.text

yuanshimingcheng_200_29 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_29 = yuanshimingcheng_200_29.text
biaotou200_29 = "此次" + name_200_29 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_29)
print('账号为：',name_200_29 + " ；上传量：",shangchuan_200_29 , ' ；播放量：' , bofang_200_29 )

worksheet.write('H144', name_200_29,geshi)
worksheet.write('I144', shangchuan_200_29,geshi)
worksheet.write('J144', bofang_200_29,geshi)

#历史好好吃
id_200_30 = '17412579530750'
url_200_30 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2615250&columnId=' + id_200_30
Browser.get(url_200_30)
time.sleep(4)

bf_200_30 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_30.text[-1].strip() == '万':
    bofang_200_30 = int(float(bf_200_30.text[:-1])*10000)
else:
    bofang_200_30 = bf_200_30.text

sc_200_30 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_30.text[-1].strip() == '集':
    shangchuan_200_30 = int(float(sc_200_30.text[:-1]))
else:
    shangchuan_200_30 = sc_200_30.text

yuanshimingcheng_200_30 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_30 = yuanshimingcheng_200_30.text
biaotou200_30 = "此次" + name_200_30 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_30)
print('账号为：',name_200_30 + " ；上传量：",shangchuan_200_30 , ' ；播放量：' , bofang_200_30 )

worksheet.write('H145', name_200_30,geshi)
worksheet.write('I145', shangchuan_200_30,geshi)
worksheet.write('J145', bofang_200_30,geshi)

#知食分子
id_200_31 = '17377147225280'
url_200_31 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2676337&columnId=' + id_200_31
Browser.get(url_200_31)
time.sleep(4)

bf_200_31 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_31.text[-1].strip() == '万':
    bofang_200_31 = int(float(bf_200_31.text[:-1])*10000)
else:
    bofang_200_31 = bf_200_31.text

sc_200_31 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_31.text[-1].strip() == '集':
    shangchuan_200_31 = int(float(sc_200_31.text[:-1]))
else:
    shangchuan_200_31 = sc_200_31.text

yuanshimingcheng_200_31 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_31 = yuanshimingcheng_200_31.text
biaotou200_31 = "此次" + name_200_31 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_31)
print('账号为：',name_200_31 + " ；上传量：",shangchuan_200_31 , ' ；播放量：' , bofang_200_31 )

worksheet.write('H146', name_200_31,geshi)
worksheet.write('I146', shangchuan_200_31,geshi)
worksheet.write('J146', bofang_200_31,geshi)

#乳腺科悄悄话
id_200_32 = '17260567653820'
url_200_32 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2628106&columnId=' + id_200_32
Browser.get(url_200_32)
time.sleep(4)

bf_200_32 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_32.text[-1].strip() == '万':
    bofang_200_32 = int(float(bf_200_32.text[:-1])*10000)
else:
    bofang_200_32 = bf_200_32.text

sc_200_32 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_32.text[-1].strip() == '集':
    shangchuan_200_32 = int(float(sc_200_32.text[:-1]))
else:
    shangchuan_200_32 = sc_200_32.text

yuanshimingcheng_200_32 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_32 = yuanshimingcheng_200_32.text
biaotou200_32 = "此次" + name_200_32 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_32)
print('账号为：',name_200_32 + " ；上传量：",shangchuan_200_32 , ' ；播放量：' , bofang_200_32 )

worksheet.write('H147', name_200_32,geshi)
worksheet.write('I147', shangchuan_200_32,geshi)
worksheet.write('J147', bofang_200_32,geshi)

#高考志愿“农情”调度
id_200_33 = '16887254077950'
url_200_33 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2179395&columnId=' + id_200_33
Browser.get(url_200_33)
time.sleep(4)

bf_200_33 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_33.text[-1].strip() == '万':
    bofang_200_33 = int(float(bf_200_33.text[:-1])*10000)
else:
    bofang_200_33 = bf_200_33.text

sc_200_33 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_33.text[-1].strip() == '集':
    shangchuan_200_33 = int(float(sc_200_33.text[:-1]))
else:
    shangchuan_200_33 = sc_200_33.text

yuanshimingcheng_200_33 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_33 = yuanshimingcheng_200_33.text
biaotou200_33 = "此次" + name_200_33 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_33)
print('账号为：',name_200_33 + " ；上传量：",shangchuan_200_33 , ' ；播放量：' , bofang_200_33 )

worksheet.write('H148', name_200_33,geshi)
worksheet.write('I148', shangchuan_200_33,geshi)
worksheet.write('J148', bofang_200_33,geshi)

#听见丰收（菁彩声版）
id_200_34 = '16710722312500'
url_200_34 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2012317&columnId=' + id_200_34
Browser.get(url_200_34)
time.sleep(4)

bf_200_34 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_34.text[-1].strip() == '万':
    bofang_200_34 = int(float(bf_200_34.text[:-1])*10000)
else:
    bofang_200_34 = bf_200_34.text

sc_200_34 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_34.text[-1].strip() == '集':
    shangchuan_200_34 = int(float(sc_200_34.text[:-1]))
else:
    shangchuan_200_34 = sc_200_34.text

yuanshimingcheng_200_34 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_34 = yuanshimingcheng_200_34.text
biaotou200_34 = "此次" + name_200_34 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_34)
print('账号为：',name_200_34 + " ；上传量：",shangchuan_200_34 , ' ；播放量：' , bofang_200_34 )

worksheet.write('H149', name_200_34,geshi)
worksheet.write('I149', shangchuan_200_34,geshi)
worksheet.write('J149', bofang_200_34,geshi)

#声动中国
id_200_35 = '16636695094350'
url_200_35 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1946400&columnId=' + id_200_35
Browser.get(url_200_35)
time.sleep(4)

bf_200_35 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_35.text[-1].strip() == '万':
    bofang_200_35 = int(float(bf_200_35.text[:-1])*10000)
else:
    bofang_200_35 = bf_200_35.text

sc_200_35 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_35.text[-1].strip() == '集':
    shangchuan_200_35 = int(float(sc_200_35.text[:-1]))
else:
    shangchuan_200_35 = sc_200_35.text

yuanshimingcheng_200_35 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_35 = yuanshimingcheng_200_35.text
biaotou200_35 = "此次" + name_200_35 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_35)
print('账号为：',name_200_35 + " ；上传量：",shangchuan_200_35 , ' ；播放量：' , bofang_200_35 )

worksheet.write('H150', name_200_35,geshi)
worksheet.write('I150', shangchuan_200_35,geshi)
worksheet.write('J150', bofang_200_35,geshi)

#两会大家说
id_200_36 = '16463784411430'
url_200_36 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1678908&columnId=' + id_200_36
Browser.get(url_200_36)
time.sleep(4)

bf_200_36 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_36.text[-1].strip() == '万':
    bofang_200_36 = int(float(bf_200_36.text[:-1])*10000)
else:
    bofang_200_36 = bf_200_36.text

sc_200_36 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_36.text[-1].strip() == '集':
    shangchuan_200_36 = int(float(sc_200_36.text[:-1]))
else:
    shangchuan_200_36 = sc_200_36.text

yuanshimingcheng_200_36 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_36 = yuanshimingcheng_200_36.text
biaotou200_36 = "此次" + name_200_36 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_36)
print('账号为：',name_200_36 + " ；上传量：",shangchuan_200_36 , ' ；播放量：' , bofang_200_36 )

worksheet.write('H151', name_200_36,geshi)
worksheet.write('I151', shangchuan_200_36,geshi)
worksheet.write('J151', bofang_200_36,geshi)

#解读2022中央一号文件
id_200_37 = '16454881667840'
url_200_37 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1671941&columnId=' + id_200_37
Browser.get(url_200_37)
time.sleep(4)

bf_200_37 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_37.text[-1].strip() == '万':
    bofang_200_37 = int(float(bf_200_37.text[:-1])*10000)
else:
    bofang_200_37 = bf_200_37.text

sc_200_37 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_37.text[-1].strip() == '集':
    shangchuan_200_37 = int(float(sc_200_37.text[:-1]))
else:
    shangchuan_200_37 = sc_200_37.text

yuanshimingcheng_200_37 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_37 = yuanshimingcheng_200_37.text
biaotou200_37 = "此次" + name_200_37 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_37)
print('账号为：',name_200_37 + " ；上传量：",shangchuan_200_37 , ' ；播放量：' , bofang_200_37 )

worksheet.write('H152', name_200_37,geshi)
worksheet.write('I152', shangchuan_200_37,geshi)
worksheet.write('J152', bofang_200_37,geshi)

#乡音闹元宵
id_200_38 = '16448185477670'
url_200_38 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1666013&columnId=' + id_200_38
Browser.get(url_200_38)
time.sleep(4)

bf_200_38 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_38.text[-1].strip() == '万':
    bofang_200_38 = int(float(bf_200_38.text[:-1])*10000)
else:
    bofang_200_38 = bf_200_38.text

sc_200_38 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_38.text[-1].strip() == '集':
    shangchuan_200_38 = int(float(sc_200_38.text[:-1]))
else:
    shangchuan_200_38 = sc_200_38.text

yuanshimingcheng_200_38 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_38 = yuanshimingcheng_200_38.text
biaotou200_38 = "此次" + name_200_38 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_38)
print('账号为：',name_200_38 + " ；上传量：",shangchuan_200_38 , ' ；播放量：' , bofang_200_38 )

worksheet.write('H153', name_200_38,geshi)
worksheet.write('I153', shangchuan_200_38,geshi)
worksheet.write('J153', bofang_200_38,geshi)

#乡聚中国年 听见乡村好日子
id_200_39 = '16425603751720'
url_200_39 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1653109&columnId=' + id_200_39
Browser.get(url_200_39)
time.sleep(4)

bf_200_39 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_39.text[-1].strip() == '万':
    bofang_200_39 = int(float(bf_200_39.text[:-1])*10000)
else:
    bofang_200_39 = bf_200_39.text

sc_200_39 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_39.text[-1].strip() == '集':
    shangchuan_200_39 = int(float(sc_200_39.text[:-1]))
else:
    shangchuan_200_39 = sc_200_39.text

yuanshimingcheng_200_39 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_39 = yuanshimingcheng_200_39.text
biaotou200_39 = "此次" + name_200_39 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_39)
print('账号为：',name_200_39 + " ；上传量：",shangchuan_200_39 , ' ；播放量：' , bofang_200_39 )

worksheet.write('H154', name_200_39,geshi)
worksheet.write('I154', shangchuan_200_39,geshi)
worksheet.write('J154', bofang_200_39,geshi)

#家“盒”年味浓
id_200_40 = '16425602654670'
url_200_40 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1653105&columnId=' + id_200_40
Browser.get(url_200_40)
time.sleep(4)

bf_200_40 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_40.text[-1].strip() == '万':
    bofang_200_40 = int(float(bf_200_40.text[:-1])*10000)
else:
    bofang_200_40 = bf_200_40.text

sc_200_40 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_40.text[-1].strip() == '集':
    shangchuan_200_40 = int(float(sc_200_40.text[:-1]))
else:
    shangchuan_200_40 = sc_200_40.text

yuanshimingcheng_200_40 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_40 = yuanshimingcheng_200_40.text
biaotou200_40 = "此次" + name_200_40 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_40)
print('账号为：',name_200_40 + " ；上传量：",shangchuan_200_40 , ' ；播放量：' , bofang_200_40 )

worksheet.write('H155', name_200_40,geshi)
worksheet.write('I155', shangchuan_200_40,geshi)
worksheet.write('J155', bofang_200_40,geshi)

#给你的心里话
id_200_41 = '16423998628180'
url_200_41 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1657957&columnId=' + id_200_41
Browser.get(url_200_41)
time.sleep(4)

bf_200_41 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_41.text[-1].strip() == '万':
    bofang_200_41 = int(float(bf_200_41.text[:-1])*10000)
else:
    bofang_200_41 = bf_200_41.text

sc_200_41 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_41.text[-1].strip() == '集':
    shangchuan_200_41 = int(float(sc_200_41.text[:-1]))
else:
    shangchuan_200_41 = sc_200_41.text

yuanshimingcheng_200_41 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_41 = yuanshimingcheng_200_41.text
biaotou200_41 = "此次" + name_200_41 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_41)
print('账号为：',name_200_41 + " ；上传量：",shangchuan_200_41 , ' ；播放量：' , bofang_200_41 )

worksheet.write('H156', name_200_41,geshi)
worksheet.write('I156', shangchuan_200_41,geshi)
worksheet.write('J156', bofang_200_41,geshi)


# 探秘千年神秘粮仓
id_200_44 = '16079357857070'
url_200_44 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=1379580&columnId=' + id_200_44
Browser.get(url_200_44)
time.sleep(4)

bf_200_44 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_44.text[-1].strip() == '万':
    bofang_200_44 = int(float(bf_200_44.text[:-1])*10000)
else:
    bofang_200_44 = bf_200_44.text

sc_200_44 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_44.text[-1].strip() == '集':
    shangchuan_200_44 = int(float(sc_200_44.text[:-1]))
else:
    shangchuan_200_44 = sc_200_44.text

yuanshimingcheng_200_44 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_44 = yuanshimingcheng_200_44.text
biaotou200_44 = "此次" + name_200_44 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_44)
print('账号为：',name_200_44 + " ；上传量：",shangchuan_200_44 , ' ；播放量：' , bofang_200_44 )

worksheet.write('H157', name_200_44,geshi)
worksheet.write('I157', shangchuan_200_44,geshi)
worksheet.write('J157', bofang_200_44,geshi)


# # 我的家乡有国宝
# id_200_45 = '17454799758520'
# url_200_45 = 'https://ytweb.radio.cn/share/albumPlay?correlateId=2704186&columnId=' + id_200_45
# Browser.get(url_200_45)
# time.sleep(4)
#
# bf_200_45 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
# if bf_200_45.text[-1].strip() == '万':
#     bofang_200_45 = int(float(bf_200_45.text[:-1])*10000)
# else:
#     bofang_200_45 = bf_200_45.text
#
# sc_200_45 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
# if sc_200_45.text[-1].strip() == '集':
#     shangchuan_200_45 = int(float(sc_200_45.text[:-1]))
# else:
#     shangchuan_200_45 = sc_200_45.text
#
# yuanshimingcheng_200_45 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
# name_200_45 = yuanshimingcheng_200_45.text
# biaotou200_45 = "此次" + name_200_45 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))
#
# print (biaotou200_44)
# print('账号为：',name_200_45 + " ；上传量：",shangchuan_200_45 , ' ；播放量：' , bofang_200_45 )
#
# worksheet.write('H158', name_200_45,geshi)
# worksheet.write('I158', shangchuan_200_45,geshi)
# worksheet.write('J158', bofang_200_45,geshi)
#专辑新增 200_50

#---------------------------------------------------------------------------------------------------------以下是回听



biaotou4_1_0 = "此次云听回放数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

worksheet.set_column('L:L',40)

worksheet.merge_range('L1:N1',biaotou4_1_0,biaotougeshi)

worksheet.write('L2', '账号名称',biaotougeshi)
worksheet.write('M2', '上传量',biaotougeshi)
worksheet.write('N2', '播放量',biaotougeshi)

#乡村振兴资讯（回听）
id_4_43 = '1396464'
url_4_43 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=244848&columnId=' + id_4_43
Browser.get(url_4_43)
time.sleep(4)

bf_4_43 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_43.text[-1].strip() == '万':
    bofang_4_43 = int(float(bf_4_43.text[:-1])*10000)
else:
    bofang_4_43 = bf_4_43.text

sc_4_43 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_43.text[-1].strip() == '集':
    shangchuan_4_43 = int(float(sc_4_43.text[:-1]))
else:
    shangchuan_4_43 = sc_4_43.text

yuanshimingcheng_4_43 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_43 = yuanshimingcheng_4_43.text + '(回听)'
biaotou4_43 = "此次" + name_4_43 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_43)
print('账号为：',name_4_43 + " ；上传量：",shangchuan_4_43 , ' ；播放量：' , bofang_4_43 )

worksheet.write('L3', name_4_43,geshi)
worksheet.write('M3', shangchuan_4_43,geshi)
worksheet.write('N3', bofang_4_43,geshi)

#健康到家（回听）
id_4_44 = '1396463'
url_4_44 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=5230480&columnId=' + id_4_44
Browser.get(url_4_44)
time.sleep(4)

bf_4_44 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_44.text[-1].strip() == '万':
    bofang_4_44 = int(float(bf_4_44.text[:-1])*10000)
else:
    bofang_4_44 = bf_4_44.text

sc_4_44 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_44.text[-1].strip() == '集':
    shangchuan_4_44 = int(float(sc_4_44.text[:-1]))
else:
    shangchuan_4_44 = sc_4_44.text

yuanshimingcheng_4_44 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_44 = yuanshimingcheng_4_44.text + '(回听)'
biaotou4_44 = "此次" + name_4_44 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_44)
print('账号为：',name_4_44 + " ；上传量：",shangchuan_4_44 , ' ；播放量：' , bofang_4_44 )

worksheet.write('L4', name_4_44,geshi)
worksheet.write('M4', shangchuan_4_44,geshi)
worksheet.write('N4', bofang_4_44,geshi)

#追梦2050（回听）
id_4_45 = '1396461'
url_4_45 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=6061002&columnId=' + id_4_45
Browser.get(url_4_45)
time.sleep(4)

bf_4_45 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_45.text[-1].strip() == '万':
    bofang_4_45 = int(float(bf_4_45.text[:-1])*10000)
else:
    bofang_4_45 = bf_4_45.text

sc_4_45 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_45.text[-1].strip() == '集':
    shangchuan_4_45 = int(float(sc_4_45.text[:-1]))
else:
    shangchuan_4_45 = sc_4_45.text

yuanshimingcheng_4_45 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_45 = yuanshimingcheng_4_45.text + '(回听)'
biaotou4_45 = "此次" + name_4_45 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_45)
print('账号为：',name_4_45 + " ；上传量：",shangchuan_4_45 , ' ；播放量：' , bofang_4_45 )

worksheet.write('L5', name_4_45,geshi)
worksheet.write('M5', shangchuan_4_45,geshi)
worksheet.write('N5', bofang_4_45,geshi)

#善治之路（回听）
id_4_46 = '1396465'
url_4_46 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=6061005&columnId=' + id_4_46
Browser.get(url_4_46)
time.sleep(4)

bf_4_46 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_46.text[-1].strip() == '万':
    bofang_4_46 = int(float(bf_4_46.text[:-1])*10000)
else:
    bofang_4_46 = bf_4_46.text

sc_4_46 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_46.text[-1].strip() == '集':
    shangchuan_4_46 = int(float(sc_4_46.text[:-1]))
else:
    shangchuan_4_46 = sc_4_46.text

yuanshimingcheng_4_46 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_46 = yuanshimingcheng_4_46.text + '(回听)'
biaotou4_46 = "此次" + name_4_46 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_46)
print('账号为：',name_4_46 + " ；上传量：",shangchuan_4_46 , ' ；播放量：' , bofang_4_46 )

worksheet.write('L6', name_4_46,geshi)
worksheet.write('M6', shangchuan_4_46,geshi)
worksheet.write('N6', bofang_4_46,geshi)

#山水乡愁（回听）
id_4_47 = '1396460'
url_4_47 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=6061012&columnId=' + id_4_47
Browser.get(url_4_47)
time.sleep(4)

bf_4_47 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_47.text[-1].strip() == '万':
    bofang_4_47 = int(float(bf_4_47.text[:-1])*10000)
else:
    bofang_4_47 = bf_4_47.text

sc_4_47 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_47.text[-1].strip() == '集':
    shangchuan_4_47 = int(float(sc_4_47.text[:-1]))
else:
    shangchuan_4_47 = sc_4_47.text

yuanshimingcheng_4_47 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_47 = yuanshimingcheng_4_47.text + '(回听)'
biaotou4_47 = "此次" + name_4_47 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_47)
print('账号为：',name_4_47 + " ；上传量：",shangchuan_4_47 , ' ；播放量：' , bofang_4_47 )

worksheet.write('L7', name_4_47,geshi)
worksheet.write('M7', shangchuan_4_47,geshi)
worksheet.write('N7', bofang_4_47,geshi)

#乡村音乐（回听）
id_4_48 = '1395826'
url_4_48 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=6061003&columnId=' + id_4_48
Browser.get(url_4_48)
time.sleep(4)

bf_4_48 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_48.text[-1].strip() == '万':
    bofang_4_48 = int(float(bf_4_48.text[:-1])*10000)
else:
    bofang_4_48 = bf_4_48.text

sc_4_48 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_48.text[-1].strip() == '集':
    shangchuan_4_48 = int(float(sc_4_48.text[:-1]))
else:
    shangchuan_4_48 = sc_4_48.text

yuanshimingcheng_4_48 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_48 = yuanshimingcheng_4_48.text + '(回听)'
biaotou4_48 = "此次" + name_4_48 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_48)
print('账号为：',name_4_48 + " ；上传量：",shangchuan_4_48 , ' ；播放量：' , bofang_4_48 )

worksheet.write('L8', name_4_48,geshi)
worksheet.write('M8', shangchuan_4_48,geshi)
worksheet.write('N8', bofang_4_48,geshi)

#绽放梨园（回听）
id_4_74 = '1396466'
url_4_74 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=4108851&columnId=' + id_4_74
Browser.get(url_4_74)
time.sleep(4)

bf_4_74 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_74.text[-1].strip() == '万':
    bofang_4_74 = int(float(bf_4_74.text[:-1])*10000)
else:
    bofang_4_74 = bf_4_74.text

sc_4_74 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_74.text[-1].strip() == '集':
    shangchuan_4_74 = int(float(sc_4_74.text[:-1]))
else:
    shangchuan_4_74 = sc_4_74.text

yuanshimingcheng_4_74 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_74 = yuanshimingcheng_4_74.text + '(回听)'
biaotou4_74 = "此次" + name_4_74 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_74)
print('账号为：',name_4_74 + " ；上传量：",shangchuan_4_74 , ' ；播放量：' , bofang_4_74 )

worksheet.write('L9', name_4_74,geshi)
worksheet.write('M9', shangchuan_4_74,geshi)
worksheet.write('N9', bofang_4_74,geshi)

#田野听书（回听）
id_4_75 = '1396459'
url_4_75 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=5967922&columnId=' + id_4_75
Browser.get(url_4_75)
time.sleep(4)

bf_4_75 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_75.text[-1].strip() == '万':
    bofang_4_75 = int(float(bf_4_75.text[:-1])*10000)
else:
    bofang_4_75 = bf_4_75.text

sc_4_75 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_75.text[-1].strip() == '集':
    shangchuan_4_75 = int(float(sc_4_75.text[:-1]))
else:
    shangchuan_4_75 = sc_4_75.text

yuanshimingcheng_4_75 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_75 = yuanshimingcheng_4_75.text + '(回听)'
biaotou4_75 = "此次" + name_4_75 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_75)
print('账号为：',name_4_75 + " ；上传量：",shangchuan_4_75 , ' ；播放量：' , bofang_4_75 )

worksheet.write('L10', name_4_75,geshi)
worksheet.write('M10', shangchuan_4_75,geshi)
worksheet.write('N10', bofang_4_75,geshi)

#田野听书（重播）
id_4_76 = '1396479'
url_4_76 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=5966956&columnId=' + id_4_76
Browser.get(url_4_76)
time.sleep(4)

bf_4_76 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_76.text[-1].strip() == '万':
    bofang_4_76 = int(float(bf_4_76.text[:-1])*10000)
else:
    bofang_4_76 = bf_4_76.text

sc_4_76 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_76.text[-1].strip() == '集':
    shangchuan_4_76 = int(float(sc_4_76.text[:-1]))
else:
    shangchuan_4_76 = sc_4_76.text

yuanshimingcheng_4_76 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_76 = yuanshimingcheng_4_76.text + '(回听)'
biaotou4_76 = "此次" + name_4_76 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_76)
print('账号为：',name_4_76 + " ；上传量：",shangchuan_4_76 , ' ；播放量：' , bofang_4_76 )

worksheet.write('L11', name_4_76,geshi)
worksheet.write('M11', shangchuan_4_76,geshi)
worksheet.write('N11', bofang_4_76,geshi)

#快乐曲艺
id_4_77 = '1396462'
url_4_77 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=4313249&columnId=' + id_4_77
Browser.get(url_4_77)
time.sleep(4)

bf_4_77 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_77.text[-1].strip() == '万':
    bofang_4_77 = int(float(bf_4_77.text[:-1])*10000)
else:
    bofang_4_77 = bf_4_77.text

sc_4_77 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_77.text[-1].strip() == '集':
    shangchuan_4_77 = int(float(sc_4_77.text[:-1]))
else:
    shangchuan_4_77 = sc_4_77.text

yuanshimingcheng_4_77 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_77 = yuanshimingcheng_4_77.text + '(回听)'
biaotou4_77 = "此次" + name_4_77 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_77)
print('账号为：',name_4_77 + " ；上传量：",shangchuan_4_77 , ' ；播放量：' , bofang_4_77 )

worksheet.write('L12', name_4_77,geshi)
worksheet.write('M12', shangchuan_4_77,geshi)
worksheet.write('N12', bofang_4_77,geshi)

#快乐曲艺（重播）
id_4_78 = '1396482'
url_4_78 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=4162552&columnId=' + id_4_78
Browser.get(url_4_78)
time.sleep(4)

bf_4_78 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_78.text[-1].strip() == '万':
    bofang_4_78 = int(float(bf_4_78.text[:-1])*10000)
else:
    bofang_4_78 = bf_4_78.text

sc_4_78 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_78.text[-1].strip() == '集':
    shangchuan_4_78 = int(float(sc_4_78.text[:-1]))
else:
    shangchuan_4_78 = sc_4_78.text

yuanshimingcheng_4_78 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_78 = yuanshimingcheng_4_78.text + '(回听)'
biaotou4_78 = "此次" + name_4_78 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_78)
print('账号为：',name_4_78 + " ；上传量：",shangchuan_4_78 , ' ；播放量：' , bofang_4_78 )

worksheet.write('L13', name_4_78,geshi)
worksheet.write('M13', shangchuan_4_78,geshi)
worksheet.write('N13', bofang_4_78,geshi)

#乡村音乐（重播）
id_4_79 = '1396470'
url_4_79 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=6648043&columnId=' + id_4_79
Browser.get(url_4_79)
time.sleep(4)

bf_4_79 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_79.text[-1].strip() == '万':
    bofang_4_79 = int(float(bf_4_79.text[:-1])*10000)
else:
    bofang_4_79 = bf_4_79.text

sc_4_79 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_79.text[-1].strip() == '集':
    shangchuan_4_79 = int(float(sc_4_79.text[:-1]))
else:
    shangchuan_4_79 = sc_4_79.text

yuanshimingcheng_4_79 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_79 = yuanshimingcheng_4_79.text + '(回听)'
biaotou4_79 = "此次" + name_4_79 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_79)
print('账号为：',name_4_79 + " ；上传量：",shangchuan_4_79 , ' ；播放量：' , bofang_4_79 )

worksheet.write('L14', name_4_79,geshi)
worksheet.write('M14', shangchuan_4_79,geshi)
worksheet.write('N14', bofang_4_79,geshi)

#山水乡愁（重播）
id_4_80 = '1396470'
url_4_80 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=6648043&columnId=' + id_4_80
Browser.get(url_4_80)
time.sleep(4)

bf_4_80 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_80.text[-1].strip() == '万':
    bofang_4_80 = int(float(bf_4_80.text[:-1])*10000)
else:
    bofang_4_80 = bf_4_80.text

sc_4_80 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_80.text[-1].strip() == '集':
    shangchuan_4_80 = int(float(sc_4_80.text[:-1]))
else:
    shangchuan_4_80 = sc_4_80.text

yuanshimingcheng_4_80 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_80 = yuanshimingcheng_4_80.text + '(回听)'
biaotou4_80 = "此次" + name_4_80 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_80)
print('账号为：',name_4_80 + " ；上传量：",shangchuan_4_80 , ' ；播放量：' , bofang_4_80 )

worksheet.write('L15', name_4_80,geshi)
worksheet.write('M15', shangchuan_4_80,geshi)
worksheet.write('N15', bofang_4_80,geshi)

#梨园乡韵（重播）
id_4_81 = '1396948'
url_4_81 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=6941507&columnId=' + id_4_81
Browser.get(url_4_81)
time.sleep(4)

bf_4_81 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_81.text[-1].strip() == '万':
    bofang_4_81 = int(float(bf_4_81.text[:-1])*10000)
else:
    bofang_4_81 = bf_4_81.text

sc_4_81 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_81.text[-1].strip() == '集':
    shangchuan_4_81 = int(float(sc_4_81.text[:-1]))
else:
    shangchuan_4_81 = sc_4_81.text

yuanshimingcheng_4_81 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_81 = yuanshimingcheng_4_81.text + '(回听)'
biaotou4_81 = "此次" + name_4_81 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_81)
print('账号为：',name_4_81 + " ；上传量：",shangchuan_4_81 , ' ；播放量：' , bofang_4_81 )

worksheet.write('L16', name_4_81,geshi)
worksheet.write('M16', shangchuan_4_81,geshi)
worksheet.write('N16', bofang_4_81,geshi)


#这里有说法（回听）
id_4_87 = '1396945'
url_4_87 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=38891363&columnId=' + id_4_87
Browser.get(url_4_87)
time.sleep(4)

bf_4_87 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_87.text[-1].strip() == '万':
    bofang_4_87 = int(float(bf_4_87.text[:-1])*10000)
else:
    bofang_4_87 = bf_4_87.text

sc_4_87 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_87.text[-1].strip() == '集':
    shangchuan_4_87 = int(float(sc_4_87.text[:-1]))
else:
    shangchuan_4_87 = sc_4_87.text

yuanshimingcheng_4_87 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_87 = yuanshimingcheng_4_87.text + '(回听)'
biaotou4_87 = "此次" + name_4_87 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_87)
print('账号为：',name_4_87 + " ；上传量：",shangchuan_4_87 , ' ；播放量：' , bofang_4_87 )

worksheet.write('L17', name_4_87,geshi)
worksheet.write('M17', shangchuan_4_87,geshi)
worksheet.write('N17', bofang_4_87,geshi)


#中国三农报道（回听）
id_4_95 = '1396942'
url_4_95 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=6993211&columnId=' + id_4_95
Browser.get(url_4_95)
time.sleep(4)

bf_4_95 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_95.text[-1].strip() == '万':
    bofang_4_95 = int(float(bf_4_95.text[:-1])*10000)
else:
    bofang_4_95 = bf_4_95.text

sc_4_95 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_95.text[-1].strip() == '集':
    shangchuan_4_95 = int(float(sc_4_95.text[:-1]))
else:
    shangchuan_4_95 = sc_4_95.text

yuanshimingcheng_4_95 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_95 = yuanshimingcheng_4_95.text + '(回听)'
biaotou4_95 = "此次" + name_4_95 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_95)
print('账号为：',name_4_95 + " ；上传量：",shangchuan_4_95 , ' ；播放量：' , bofang_4_95 )

worksheet.write('L18', name_4_95,geshi)
worksheet.write('M18', shangchuan_4_95,geshi)
worksheet.write('N18', bofang_4_95,geshi)


#田园新主张（回听）
id_4_130 = '1396943'
url_4_130 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=325486891&columnId=' + id_4_130
Browser.get(url_4_130)
time.sleep(4)

bf_4_130 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_4_130.text[-1].strip() == '万':
    bofang_4_130 = int(float(bf_4_130.text[:-1])*10000)
else:
    bofang_4_130 = bf_4_130.text

sc_4_130 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_4_130.text[-1].strip() == '集':
    shangchuan_4_130 = int(float(sc_4_130.text[:-1]))
else:
    shangchuan_4_130 = sc_4_130.text

yuanshimingcheng_4_130 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_4_130 = yuanshimingcheng_4_130.text + '(回听)'
biaotou4_130 = "此次" + name_4_130 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_130)
print('账号为：',name_4_130 + " ；上传量：",shangchuan_4_130 , ' ；播放量：' , bofang_4_130 )

worksheet.write('L19', name_4_130,geshi)
worksheet.write('M19', shangchuan_4_130,geshi)
worksheet.write('N19', bofang_4_130,geshi)


#听见乡村（回听）
id_5_130 = '1439136'
url_5_130 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=963935807&columnId=' + id_5_130
Browser.get(url_5_130)
time.sleep(4)

bf_5_130 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_5_130.text[-1].strip() == '万':
    bofang_5_130 = int(float(bf_5_130.text[:-1])*10000)
else:
    bofang_5_130 = bf_5_130.text

sc_5_130 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_5_130.text[-1].strip() == '集':
    shangchuan_5_130 = int(float(sc_5_130.text[:-1]))
else:
    shangchuan_5_130 = sc_5_130.text

yuanshimingcheng_5_130 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_5_130 = yuanshimingcheng_5_130.text + '(回听)'
biaotou4_130 = "此次" + name_5_130 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_130)
print('账号为：',name_5_130 + " ；上传量：",shangchuan_5_130 , ' ；播放量：' , bofang_5_130 )

worksheet.write('L20', name_5_130,geshi)
worksheet.write('M20', shangchuan_5_130,geshi)
worksheet.write('N20', bofang_5_130,geshi)

#听见乡村（回听）重播
id_5_131 = '1439137'
url_5_131 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=963949979&columnId=' + id_5_131
Browser.get(url_5_131)
time.sleep(4)

bf_5_131 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_5_131.text[-1].strip() == '万':
    bofang_5_131 = int(float(bf_5_131.text[:-1])*10000)
else:
    bofang_5_131 = bf_5_131.text

sc_5_131 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_5_131.text[-1].strip() == '集':
    shangchuan_5_131 = int(float(sc_5_131.text[:-1]))
else:
    shangchuan_5_131 = sc_5_131.text

yuanshimingcheng_5_131 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_5_131 = yuanshimingcheng_5_131.text + '(回听)'
biaotou4_130 = "此次" + name_5_131 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou4_130)
print('账号为：',name_5_131 + " ；上传量：",shangchuan_5_131 , ' ；播放量：' , bofang_5_131 )

worksheet.write('L21', name_5_131,geshi)
worksheet.write('M21', shangchuan_5_131,geshi)
worksheet.write('N21', bofang_5_131,geshi)

#田园新主张(重播)(回听)
id_200_11 = '1396946'
url_200_11 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=6923060&columnId=' + id_200_11
Browser.get(url_200_11)
time.sleep(4)

bf_200_11 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_11.text[-1].strip() == '万':
    bofang_200_11 = int(float(bf_200_11.text[:-1])*10000)
else:
    bofang_200_11 = bf_200_11.text

sc_200_11 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_11.text[-1].strip() == '集':
    shangchuan_200_11 = int(float(sc_200_11.text[:-1]))
else:
    shangchuan_200_11 = sc_200_11.text

yuanshimingcheng_200_11 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_11 = yuanshimingcheng_200_11.text + '(回听)'
biaotou200_11 = "此次" + name_200_11 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_11)
print('账号为：',name_200_11 + " ；上传量：",shangchuan_200_11 , ' ；播放量：' , bofang_200_11 )

worksheet.write('L22', name_200_11,geshi)
worksheet.write('M22', shangchuan_200_11,geshi)
worksheet.write('N22', bofang_200_11,geshi)

#田园新主张(精编版)(回听)
id_200_12 = '1439138'
url_200_12 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=962011610&columnId=' + id_200_12
Browser.get(url_200_12)
time.sleep(4)

bf_200_12 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_12.text[-1].strip() == '万':
    bofang_200_12 = int(float(bf_200_12.text[:-1])*10000)
else:
    bofang_200_12 = bf_200_12.text

sc_200_12 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_12.text[-1].strip() == '集':
    shangchuan_200_12 = int(float(sc_200_12.text[:-1]))
else:
    shangchuan_200_12 = sc_200_12.text

yuanshimingcheng_200_12 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_12 = yuanshimingcheng_200_12.text + '(回听)'
biaotou200_12 = "此次" + name_200_12 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_12)
print('账号为：',name_200_12 + " ；上传量：",shangchuan_200_12 , ' ；播放量：' , bofang_200_12 )

worksheet.write('L23', name_200_12,geshi)
worksheet.write('M23', shangchuan_200_12,geshi)
worksheet.write('N23', bofang_200_12,geshi)

#这里有说法(重播)(回听)
id_200_13 = '1396950'
url_200_13 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=6942085&columnId=' + id_200_13
Browser.get(url_200_13)
time.sleep(4)

bf_200_13 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_13.text[-1].strip() == '万':
    bofang_200_13 = int(float(bf_200_13.text[:-1])*10000)
else:
    bofang_200_13 = bf_200_13.text

sc_200_13 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_13.text[-1].strip() == '集':
    shangchuan_200_13 = int(float(sc_200_13.text[:-1]))
else:
    shangchuan_200_13 = sc_200_13.text

yuanshimingcheng_200_13 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_13 = yuanshimingcheng_200_13.text + '(回听)'
biaotou200_13 = "此次" + name_200_13 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_13)
print('账号为：',name_200_13 + " ；上传量：",shangchuan_200_13 , ' ；播放量：' , bofang_200_13 )

worksheet.write('L24', name_200_13,geshi)
worksheet.write('M24', shangchuan_200_13,geshi)
worksheet.write('N24', bofang_200_13,geshi)

#健康到家(重播)(回听)
id_200_14 = '1396951'
url_200_14 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=6923046&columnId=' + id_200_14
Browser.get(url_200_14)
time.sleep(4)

bf_200_14 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_14.text[-1].strip() == '万':
    bofang_200_14 = int(float(bf_200_14.text[:-1])*10000)
else:
    bofang_200_14 = bf_200_14.text

sc_200_14 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_14.text[-1].strip() == '集':
    shangchuan_200_14 = int(float(sc_200_14.text[:-1]))
else:
    shangchuan_200_14 = sc_200_14.text

yuanshimingcheng_200_14 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_14 = yuanshimingcheng_200_14.text + '(回听)'
biaotou200_14 = "此次" + name_200_14 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_14)
print('账号为：',name_200_14 + " ；上传量：",shangchuan_200_14 , ' ；播放量：' , bofang_200_14 )

worksheet.write('L25', name_200_14,geshi)
worksheet.write('M25', shangchuan_200_14,geshi)
worksheet.write('N25', bofang_200_14,geshi)




#振兴新力量
id_200_30 = '1451667'
url_200_30 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=977686971&columnId=' + id_200_30
Browser.get(url_200_30)
time.sleep(4)

bf_200_30 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_30.text[-1].strip() == '万':
    bofang_200_30 = int(float(bf_200_30.text[:-1])*10000)
else:
    bofang_200_30 = bf_200_30.text

sc_200_30 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_30.text[-1].strip() == '集':
    shangchuan_200_30 = int(float(sc_200_30.text[:-1]))
else:
    shangchuan_200_30 = sc_200_30.text

yuanshimingcheng_200_30 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_30 = yuanshimingcheng_200_30.text
biaotou200_30 = "此次" + name_200_30 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_30)
print('账号为：',name_200_30 + " ；上传量：",shangchuan_200_30 , ' ；播放量：' , bofang_200_30 )

worksheet.write('L26', name_200_30,geshi)
worksheet.write('M26', shangchuan_200_30,geshi)
worksheet.write('N26', bofang_200_30,geshi)


#振兴新力量（重播）
id_200_31 = '1451670'
url_200_31 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=977686954&columnId=' + id_200_31
Browser.get(url_200_31)
time.sleep(4)

bf_200_31 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_31.text[-1].strip() == '万':
    bofang_200_31 = int(float(bf_200_31.text[:-1])*10000)
else:
    bofang_200_31 = bf_200_31.text

sc_200_31 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_31.text[-1].strip() == '集':
    shangchuan_200_31 = int(float(sc_200_31.text[:-1]))
else:
    shangchuan_200_31 = sc_200_31.text

yuanshimingcheng_200_31 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_31 = yuanshimingcheng_200_31.text
biaotou200_31 = "此次" + name_200_31 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_31)
print('账号为：',name_200_31 + " ；上传量：",shangchuan_200_31 , ' ；播放量：' , bofang_200_31 )

worksheet.write('L27', name_200_31,geshi)
worksheet.write('M27', shangchuan_200_31,geshi)
worksheet.write('N27', bofang_200_31,geshi)


#山水之间
id_200_32 = '1451668'
url_200_32 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=977686926&columnId=' + id_200_32
Browser.get(url_200_32)
time.sleep(4)

bf_200_32 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_32.text[-1].strip() == '万':
    bofang_200_32 = int(float(bf_200_32.text[:-1])*10000)
else:
    bofang_200_32 = bf_200_32.text

sc_200_32 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_32.text[-1].strip() == '集':
    shangchuan_200_32 = int(float(sc_200_32.text[:-1]))
else:
    shangchuan_200_32 = sc_200_32.text

yuanshimingcheng_200_32 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_32 = yuanshimingcheng_200_32.text
biaotou200_32 = "此次" + name_200_32 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_32)
print('账号为：',name_200_32 + " ；上传量：",shangchuan_200_32 , ' ；播放量：' , bofang_200_32 )

worksheet.write('L28', name_200_32,geshi)
worksheet.write('M28', shangchuan_200_32,geshi)
worksheet.write('N28', bofang_200_32,geshi)


#山水之间（重播）
id_200_33 = '1451672'
url_200_33 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=977686967&columnId=' + id_200_33
Browser.get(url_200_33)
time.sleep(4)

bf_200_33 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_33.text[-1].strip() == '万':
    bofang_200_33 = int(float(bf_200_33.text[:-1])*10000)
else:
    bofang_200_33 = bf_200_33.text

sc_200_33 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_33.text[-1].strip() == '集':
    shangchuan_200_33 = int(float(sc_200_33.text[:-1]))
else:
    shangchuan_200_33 = sc_200_33.text

yuanshimingcheng_200_33 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_33 = yuanshimingcheng_200_33.text
biaotou200_33 = "此次" + name_200_33 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_33)
print('账号为：',name_200_33 + " ；上传量：",shangchuan_200_33 , ' ；播放量：' , bofang_200_33 )

worksheet.write('L29', name_200_33,geshi)
worksheet.write('M29', shangchuan_200_33,geshi)
worksheet.write('N29', bofang_200_33,geshi)


#话农点经
id_200_34 = '1451669'
url_200_34 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=977686950&columnId=' + id_200_34
Browser.get(url_200_34)
time.sleep(4)

bf_200_34 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_34.text[-1].strip() == '万':
    bofang_200_34 = int(float(bf_200_34.text[:-1])*10000)
else:
    bofang_200_34 = bf_200_34.text

sc_200_34 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_34.text[-1].strip() == '集':
    shangchuan_200_34 = int(float(sc_200_34.text[:-1]))
else:
    shangchuan_200_34 = sc_200_34.text

yuanshimingcheng_200_34 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_34 = yuanshimingcheng_200_34.text
biaotou200_34 = "此次" + name_200_34 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_34)
print('账号为：',name_200_34 + " ；上传量：",shangchuan_200_34 , ' ；播放量：' , bofang_200_34 )

worksheet.write('L30', name_200_34,geshi)
worksheet.write('M30', shangchuan_200_34,geshi)
worksheet.write('N30', bofang_200_34,geshi)


#话农点经（重播）
id_200_35 = '1451676'
url_200_35 = 'https://ytweb.radio.cn/share/columnPlay?correlateId=977686940&columnId=' + id_200_35
Browser.get(url_200_35)
time.sleep(4)

bf_200_35 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
if bf_200_35.text[-1].strip() == '万':
    bofang_200_35 = int(float(bf_200_35.text[:-1])*10000)
else:
    bofang_200_35 = bf_200_35.text

sc_200_35 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
if sc_200_35.text[-1].strip() == '集':
    shangchuan_200_35 = int(float(sc_200_35.text[:-1]))
else:
    shangchuan_200_35 = sc_200_35.text

yuanshimingcheng_200_35 = Browser.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/div[1]/div[1]/div[1]')
name_200_35 = yuanshimingcheng_200_35.text
biaotou200_35 = "此次" + name_200_35 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))

print (biaotou200_35)
print('账号为：',name_200_35 + " ；上传量：",shangchuan_200_35 , ' ；播放量：' , bofang_200_35 )

worksheet.write('L31', name_200_35,geshi)
worksheet.write('M31', shangchuan_200_35,geshi)
worksheet.write('N31', bofang_200_35,geshi)

#回听新增 300_1









# 央视账号添加！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！
# -------------------------------------------------------------------------------------------------------------------------
# 使用说明：账号请从此处开始添加  央视频编号规则：从id_100_1开始使用
#
# 央视频账号代码示例
# -----------------
# 账号名称（央视频）
# id_100_1 = '新的身份证号'
# url_100_1 = 'https://m.yangshipin.cn/user?cpid=' + id_100_1
# Browser.get(url_100_1)
# time.sleep(4)
# fs_100_1 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[1]/span[1]')
# if fs_100_1.text[-1].strip() == '万':
#     fensi_100_1 = int(float(fs_100_1.text[:-1])*10000)
# else:
#     fensi_100_1 = fs_100_1.text
#
# bf_100_1 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[2]/span[1]')
# if bf_100_1.text[-1].strip() == '万':
#     bofang_100_1 = int(float(bf_100_1.text[:-1])*10000)
# else:
#     bofang_100_1 = bf_100_1.text
#
# sc_100_1 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[4]/div[3]/span[1]')
# if sc_100_1.text[-1].strip() == '万':
#     shangchuan_100_1 = int(float(sc_100_1.text[:-1])*10000)
# else:
#     shangchuan_100_1 = sc_100_1.text
#
# name_100_1_0 = Browser.find_element_by_xpath('/html/body/div/div[1]/div/section/div/div[2]/div[1]/div[2]')
# name_100_1 = name_100_1_0.text + '(央视频)'
# biaotou6_20 = "此次" + name_100_1 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))
#
# print (biaotou6_20)
# print('账号为：',name_100_1 + " ；上传量：",shangchuan_100_1 , ' ；播放量：' , bofang_100_1 , " ；粉丝量：" ,fensi_100_1)
#
# worksheet.write('A表格位置', name_100_1,geshi)
# worksheet.write('B表格位置', shangchuan_100_1,geshi)
# worksheet.write('C表格位置', bofang_100_1,geshi)
# worksheet.write('D表格位置', fensi_100_1,geshi)
#
# worksheet.write('E表格位置', '部门名称',geshi)


# -------------------------------------------------------------------------------------------------------------------------
# 央视频！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！









# ----------------------------------------------------------------------------------------------------------------------------
# 云听账号代码示例
# ----------------------------------------------------------------------------------------------------------------------------
# 使用说明：账号请从此处开始添加  央视频编号规则：从id_200_1开始使用

# 
# #账号名称（云听）
# id_200_1 = '身份证号码'
# url_200_1 = '身份证号码' + id_200_1
# Browser.get(url_200_1)
# time.sleep(4)
# 
# bf_200_1 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/span')
# if bf_200_1.text[-1].strip() == '万':
#     bofang_200_1 = int(float(bf_200_1.text[:-1])*10000)
# else:
#     bofang_200_1 = bf_200_1.text
# 
# sc_200_1 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[2]/div[2]/span')
# if sc_200_1.text[-1].strip() == '集':
#     shangchuan_200_1 = int(float(sc_200_1.text[:-1]))
# else:
#     shangchuan_200_1 = sc_200_1.text
# 
# yuanshimingcheng_200_1 = Browser.find_element_by_xpath('/html/body/div[1]/div/div[2]/div[2]/div[1]/div[1]/div[1]')
# name_200_1 = yuanshimingcheng_200_1.text
# biaotou4_129 = "此次" + name_200_1 + "数据抓取时间： " + time.strftime('%Y.%m.%d %H:%M:%S ',time.localtime(time.time()))
# 
# print (biaotou4_129)
# print('账号为：',name_200_1 + " ；上传量：",shangchuan_200_1 , ' ；播放量：' , bofang_200_1 )
# 
# worksheet.write('A表格位置', name_200_1,geshi)
# worksheet.write('B表格位置', shangchuan_200_1,geshi)
# worksheet.write('C表格位置', bofang_200_1,geshi)
# 
# worksheet.write('E表格位置', '部门',geshi)

# -------------------------------------------------------------------------------------------------------------------------
# 云听！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！




















# style_end = workbook.add_format({
#     'font_size': 30,
#     "fg_color": "orange",
#     "bold": 1,
#     "align": "center",
#     "valign": "vcenter",
#     "font_color": "white"
# })
#
# worksheet.set_row(60, 60)
# worksheet.merge_range('A61:N61','此数据为自动生成，请及时备份',style_end)

workbook.close()

Browser.close()

time.sleep(4)
sys.exit()
