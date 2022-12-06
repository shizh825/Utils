import win32api

path = "D:\code\\"
# 用word打开文件 word
win32api.ShellExecute(0, 'open', 'WINWORD.EXE', path + "test.docx", '' , 1)