import os
from win32com import client
import docx

def doc2docx(path):     #转化函数
    w = client.Dispatch('Word.Application')  #word接口
    doc = w.Documents.Open(path)  #打开
    newpath = os.path.splitext(path)[0] + '.docx'
    doc.SaveAs(newpath, 12, False, "", True, "", False, False, False, False)  #转化
    doc.Close()  #关闭
    w.Quit()
    os.remove(path)
    return newpath

g = os.walk(r"C:\Users\risen\Desktop\FIN\GISSA大\文书") #浏览
for path, dir_list, file_list in g:  #循环
    for file_name in file_list:
      try:
        a = os.path.join(path, file_name) #获得文件名
        print(a)
        doc2docx(a)  #转化
      except:
        print("oops")
        continue

