import tkinter
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
import xlrd
import xlwt
import os
import copy


def setBgColorStyle(color_name):
    style = xlwt.XFStyle()
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = xlwt.Style.colour_map[color_name]
    style.pattern = pattern
    return style


style_light_orange = setBgColorStyle("light_orange")
style_light_green = setBgColorStyle("light_green")


def process(filename):
    # filename = 'resources/template_test.xlsx'
    sheet_name = 'Sheet1'
    excel_file_path = os.path.join(os.getcwd(), filename)
    xl_data = xlrd.open_workbook(excel_file_path)

    # 读取需要的数据
    sheet_1 = xl_data.sheet_by_name(sheet_name)
    rows = sheet_1.nrows
    needed_content_by_row = []
    for row in range(rows):
        row_content = sheet_1.row_values(row)
        new_row_content = []
        for col, cell_content in enumerate(row_content):
            # ['受益公司代码', 'LOB', '报销类型', '发票金额', '税额', '进项税科目', '日记账名', '摘要', '摘要2', '供应商编码', '项目编码']
            # 数字为所在列的index
            if col in [3, 6, 7, 22, 23, 25, 27, 28, 29, 30, 31]:
                new_row_content.append(row_content[col])
        needed_content_by_row.append(new_row_content)

    # 添加映射需要的行
    total_content_by_row = []
    for row, row_content in enumerate(needed_content_by_row):
        if row == 0:
            copy_and_append(total_content_by_row, row_content, '科目名编码')
        if '采购分包' in row_content:
            copy_and_append(total_content_by_row, row_content, '22020101')
            copy_and_append(total_content_by_row, row_content, '22020201')
            copy_and_append(total_content_by_row, row_content, row_content[5])
            total_content_by_row.append(['', '', '', '', '', '', '', '', '', '', '', ''])
        elif '销售分包' in row_content:
            copy_and_append(total_content_by_row, row_content, '64010501')
            copy_and_append(total_content_by_row, row_content, '22020102')
            copy_and_append(total_content_by_row, row_content, row_content[5])
            total_content_by_row.append(['', '', '', '', '', '', '', '', '', '', '', ''])
        elif '交付分包' in row_content:
            copy_and_append(total_content_by_row, row_content, '22020103')
            copy_and_append(total_content_by_row, row_content, '22020203')
            copy_and_append(total_content_by_row, row_content, row_content[5])
            copy_and_append(total_content_by_row, row_content, '64010501')
            total_content_by_row.append(['', '', '', '', '', '', '', '', '', '', '', ''])

    # 转置content
    total_content_by_col = transpose_content(total_content_by_row)

    # 写sheet
    result_excel = xlwt.Workbook('utf-8')
    result_sheet = result_excel.add_sheet('sheet1')

    new_sheet_col_title = ['公司段', '部门段', '科目段', '子目段', 'LOB段', '项目段', '内部往来', '往来段', '辅助核算', '关联LOB', '预留段3', '借方', '贷方', '批名', '日记账名称', '行说明']
    for row, row_content in enumerate(total_content_by_row):
        for col, new_title in enumerate(new_sheet_col_title):
            if 0 == row:
                result_sheet.write(row, col, new_title, style=style_light_orange)
            else:
                col_content_company = get_col_content(total_content_by_col, '受益公司代码')
                if new_title == '公司段':
                    write_data_to_excel(result_sheet, row, col, col_content_company[row])
                if new_title == '部门段':
                    if col_content_company[row] == '':
                        write_data_to_excel(result_sheet, row, col, col_content_company[row])
                    else:
                        write_data_to_excel(result_sheet, row, col, 0)
                if new_title == '科目段':
                    col_content = get_col_content(total_content_by_col, '科目名编码')
                    write_data_to_excel(result_sheet, row, col, col_content[row])
                # ----------------------------------------------------------------------------------------------
                if new_title == '子目段':
                    if '采购分包' in row_content:
                        if '22020101' in row_content or '22020201' in row_content:
                            write_data_to_excel(result_sheet, row, col, 0)
                        else:
                            write_data_to_excel(result_sheet, row, col, 'G1351003')
                    elif '销售分包' in row_content:
                        if '22020102' in row_content:
                            write_data_to_excel(result_sheet, row, col, 0)
                        else:
                            write_data_to_excel(result_sheet, row, col, 'G1326091')
                    elif '交付分包' in row_content:
                        if '22020103' in row_content or '22020203' in row_content:
                            write_data_to_excel(result_sheet, row, col, 0)
                        else:
                            write_data_to_excel(result_sheet, row, col, 'G1325002')
                # ----------------------------------------------------------------------------------------------
                if new_title == 'LOB段':
                    col_content = get_col_content(total_content_by_col, 'LOB')
                    write_data_to_excel(result_sheet, row, col, col_content[row])
                if new_title == '项目段':
                    col_content = get_col_content(total_content_by_col, '项目编码')
                    write_data_to_excel(result_sheet, row, col, col_content[row])
                if new_title == '内部往来':
                    if col_content_company[row] == '':
                        write_data_to_excel(result_sheet, row, col, col_content_company[row])
                    else:
                        write_data_to_excel(result_sheet, row, col, 0)
                if new_title == '往来段':
                    col_content = get_col_content(total_content_by_col, '供应商编码')
                    write_data_to_excel(result_sheet, row, col, col_content[row])
                if new_title == '辅助核算':
                    if col_content_company[row] == '':
                        write_data_to_excel(result_sheet, row, col, col_content_company[row])
                    else:
                        write_data_to_excel(result_sheet, row, col, 0)
                if new_title == '关联LOB':
                    if col_content_company[row] == '':
                        write_data_to_excel(result_sheet, row, col, col_content_company[row])
                    else:
                        write_data_to_excel(result_sheet, row, col, 0)
                if new_title == '预留段3':
                    if col_content_company[row] == '':
                        write_data_to_excel(result_sheet, row, col, col_content_company[row])
                    else:
                        write_data_to_excel(result_sheet, row, col, 0)
                # ----------------------------------------------------------------------------------------------
                if new_title == '借方':
                    col_content_1 = get_col_content(total_content_by_col, '税额')
                    col_content_2 = get_col_content(total_content_by_col, '发票金额')
                    if '采购分包' in row_content:
                        if '22020201' in row_content:
                            write_data_to_excel(result_sheet, row, col, col_content_2[row] - col_content_1[row])
                        elif '22020101' in row_content:
                            write_data_to_excel(result_sheet, row, col, '')
                        else:
                            write_data_to_excel(result_sheet, row, col, col_content_1[row])
                    elif '销售分包' in row_content:
                        if '64010501' in row_content:
                            write_data_to_excel(result_sheet, row, col, col_content_2[row] - col_content_1[row])
                        elif '22020102' in row_content:
                            write_data_to_excel(result_sheet, row, col, '')
                        else:
                            write_data_to_excel(result_sheet, row, col, col_content_1[row])
                    elif '交付分包' in row_content:
                        if '22020203' in row_content:
                            write_data_to_excel(result_sheet, row, col, col_content_2[row])
                        elif '64010501' in row_content:
                            write_data_to_excel(result_sheet, row, col, -float(col_content_1[row]))
                        elif '22020103' in row_content:
                            write_data_to_excel(result_sheet, row, col, '')
                        else:
                            write_data_to_excel(result_sheet, row, col, col_content_1[row])
                # ----------------------------------------------------------------------------------------------
                if new_title == '贷方':
                    col_content = get_col_content(total_content_by_col, '发票金额')
                    if '采购分包' in row_content:
                        if '22020101' in row_content:
                            write_data_to_excel(result_sheet, row, col, col_content[row])
                        else:
                            write_data_to_excel(result_sheet, row, col, '')
                    elif '销售分包' in row_content:
                        if '22020102' in row_content:
                            write_data_to_excel(result_sheet, row, col, col_content[row])
                        else:
                            write_data_to_excel(result_sheet, row, col, '')
                    elif '交付分包' in row_content:
                        if '22020103' in row_content:
                            write_data_to_excel(result_sheet, row, col, col_content[row])
                        else:
                            write_data_to_excel(result_sheet, row, col, '')
                # ----------------------------------------------------------------------------------------------
                if new_title == '批名':
                    if col_content_company[row] == '':
                        write_data_to_excel(result_sheet, row, col, col_content_company[row])
                    else:
                        write_data_to_excel(result_sheet, row, col, '-')
                if new_title == '日记账名称':
                    col_content = get_col_content(total_content_by_col, '日记账名称')
                    write_data_to_excel(result_sheet, row, col, col_content[row])
                # ----------------------------------------------------------------------------------------------
                if new_title == '行说明':
                    col_content_1 = get_col_content(total_content_by_col, '摘要')
                    if '采购分包' in row_content or '销售分包' in row_content:
                        write_data_to_excel(result_sheet, row, col, col_content_1[row])
                    else:
                        if '22020103' in row_content or '22020203' in row_content:
                            write_data_to_excel(result_sheet, row, col, col_content_1[row])
                        else:
                            col_content_2 = get_col_content(total_content_by_col, '摘要2')
                            write_data_to_excel(result_sheet, row, col, col_content_2[row])
                # ----------------------------------------------------------------------------------------------
    # 保存到文件
    result_filename = filename[:-5] + '_result.xlsx'
    result_excel.save(result_filename)
    print('========================= done ==============================')
    return result_filename

def get_col_content(content_by_col, title_name):
    ret_col_content = []
    for col, col_content in enumerate(content_by_col):
        old_title = col_content[0]
        if old_title == title_name:
            ret_col_content = col_content
            break
    return ret_col_content


def transpose_content(total_content_by_row):
    transposed = []
    for col in range(len(total_content_by_row[0])):
        col_content = []
        for row_content in total_content_by_row:
            col_content.append(row_content[col])
        transposed.append(col_content)
    return transposed


def copy_and_append(total_content_by_row, row_content, subject_name):
    tmp_content = copy.deepcopy(row_content)
    tmp_content.append(subject_name)
    total_content_by_row.append(tmp_content)


def write_data_to_excel(sheet, r, c, d):
    if d == '':
        sheet.write(r, c, d)
    else:
        sheet.write(r, c, d, style=style_light_green)


def choose_file_and_process():
    filename = filedialog.askopenfilename(initialdir=os.getcwd())
    if filename != '':
        print(filename)
        result_filename = process(filename)
        show_msg = '处理成功!\n请在指定路径查看文件：\n' + result_filename
        messagebox.showinfo('处理结果', show_msg)
    else:
        print('choose no file.')


def on_enter(btn):
    btn['background'] = 'green'

def on_leave(btn):
    btn['background'] = 'SystemButtonFace'


if __name__ == "__main__":
    # 初始化窗口
    main_window = tkinter.Tk()
    main_window.title('分包入账模板--批处理程序')
    main_window.minsize(550, 350)
    main_window.maxsize(550, 350)
    main_window.geometry('550x350')
    main_window.configure(background='#eeeeee')
    # 添加说明
    ttk.Label(main_window, text='\n\n\n\n\n使用本程序需要安装python3.6.4工具', anchor=CENTER, background='#eeeeee', font=('楷体', 16)).grid(column=0, row=1)
    ttk.Label(main_window, text='      下载地址：https://www.python.org/ftp/python/3.6.4/python-3.6.4-amd64.exe\n\n\n', anchor=CENTER, background='#eeeeee', font=('楷体', 14)).grid(column=0, row=2)
    choose_btn = ttk.Button(main_window, text='点击选择模板文件', width=15, command=choose_file_and_process)
    choose_btn.grid(column=0, row=3)
    quit_btn = ttk.Button(main_window, text='退出程序', width=15, command=main_window.destroy)
    quit_btn.grid(column=0, row=4)
    # 进入消息循环
    main_window.mainloop()