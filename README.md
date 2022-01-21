# process_entry_account
批处理分包入账模板数据，从excel文件读取

## 环境安装
使用本程序需要安装python3.6.4
主要使用环境为windows系统，所以仅列出win版本的python安装包
下载地址：
https://www.python.org/ftp/python/3.6.4/python-3.6.4-amd64.exe

## 安装依赖库
pip install -r requirements.txt
## 安装pyinstaller用于编译可执行文件
pip install pyinstaller==4.8

## 编译可执行文件命令
pyinstaller -F -w process_excel.py

在/dist/目录下可以得到exe执行程序和命令行执行程序
将exe复制到需要保存的位置就可以正常使用，建议创建快捷方式到桌面

## 注意
打开文件默认位置为可执行程序所处的位置，需要查找输入文件，输出文件为输入文件的同目录

