# process_entry_account
批处理分包入账模板数据，从excel文件读取

## 安装依赖库
pip install -r requirements.txt

## 安装pyinstaller用于编译可执行文件
pip install pyinstaller==4.8

## 编译可执行文件命令
pyinstaller -F -w process_excel.py

在/dist/目录下可以得到macOS应用XXX.app和命令行执行程序，将XXX.app移到系统应用程序就可以正常使用
