import time
import pyautogui
import pandas as pd

#运行后请立即切换到chrome浏览器
def input_id(url): #自动下载
    pyautogui.moveTo(180, 60, duration=0.2)  #定位到某一点
    pyautogui.click(button='left') #点击
    time.sleep(0.5)
    pyautogui.typewrite(url,0.01) #输入字符，0.01表示输入每个字符间隔的时间
    time.sleep(0.5)
    pyautogui.press("enter")  #点击确定

time.sleep(1)
a=pd.read_csv('Link-to-verdict/泸州.csv')   #读取文档
row=a.shape[0]                 #循环获取url
for i in range(row):
    w=a.iloc[i,8]
    w=w.replace("website/wenshu/181107ANFZ0BXSK4/index.html", "down/one", 1)  #替换
    input_id(w)



