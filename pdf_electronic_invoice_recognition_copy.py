#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Author  : Tai cu Shi ba
# Copyright (C) Tai cu Shi ba, All Rights Reserved
# @File    : pdf_electronic_invoice_recognition.py
# @IDE     : PyCharm
# -*- coding: utf-8 -*-

# Author:黄成
# Date:2023/10/11
# Function: 一键式实现电子发票识别
# 注意事项: 被识别发票文件需要解密

import pdfplumber
import re
import os
import xlwt
import sys
import datetime
from collections import defaultdict
# 异常开票纳税识别号
invalid_company_ids = [
'91440300359382172R',
'91440300MA5F9JC74N',
'92440300MA5G9MUC9M',
'92440300MA5G9MNE0F',
'92440300MA5G9MMT7U',
'92440300MA5G9QF73F',
'92440300MA5G9TLJ85',
'91440300MA5FKFMBX3',
'91440300MA5FT2LM7U',
'91440300568505124A',
'91440300MA5G2W834G',
'91440300MA5G0D0G7X',
'91440300MA5GAJ7W37',
'91440300MA5FARFG4B',
'91440300596771171E',
'92440300MA5G9YJ387',
'91440300MA5DA9CY1N',
'92440300MA5F40T737',
'91440300MA5H1J5C9T',
'91110114MA04FJXF7A',
'91440300MA5CV8XN57',
'91310110MA1G99QA4Q',
'91440300MA5GUT6524',
'91440300MA5CTUAT1V',
'91440300MA5CTXCM05',
'91440300MA5EF JUP8L',
'91440300MA5F0KF938',
'91510104MA681BEA19'
]

# 创建工作簿
wb = xlwt.Workbook()
# 创建表单
sh = wb.add_sheet('发票信息')
info_titles = ['发票名称','发票代码', '发票号码','开票日期','卖方公司','卖方公司纳税人识别号','收款人','复核','开票人','购方公司','购方公司纳税人识别号','金额','检验结果']
# 写表头
for i in range(len(info_titles)):
    sh.write(0,i,info_titles[i])
def re_text(bt, text):
    m1 = re.search(bt, text)
    if m1 is not None:
        return re_block(m1[0])

def re_block(text):
    return text.replace(' ', '').replace('　', '').replace('）', '').replace(')', '').replace('：', ':')

# 购方纳税人识别号检查
def verify_buycompany_ids(inovice_filename, buytax_num):
    if buytax_num[7:] != '9144030031977063XH':
        print(f"{inovice_filename}公司纳税人识别号错误")
        return False
    return True

# 卖方纳税人识别号检查
def verify_sellcompany_ids(inovice_filename, selltax_num):
    if selltax_num[7:] in invalid_company_ids:
        print(f"{inovice_filename}开票方纳税识别号不合规")
        return False
    return True

# 购方公司名称检查
def verify_company_name(inovice_filename, buy_company_name):
    if buy_company_name[3:] != 'XXX公司':
        print(f"{inovice_filename}公司名称错误")
        return False
    return True

# 收集发票错误信息
chk_errors = defaultdict(list)
# 开票人必须填写，复核人和收款人可不填写，复核人不可以与开票人为一人
def verify_people_info(inovice_filename, drawer, reviewer, payee):
    if not drawer or not re.findall(re.compile(r'[\u4e00-\u9fa5]+'),drawer):
        # 开票人必须填写
        chk_errors[inovice_filename].append(" |开票人信息有误| ")
    # 复核人和开票人为同一人
    if reviewer == drawer:
        chk_errors[inovice_filename].append(" |开票人和复核人不可以为同一人| ")
    if  reviewer and not re.findall(re.compile(r'[\u4e00-\u9fa5]+'),reviewer):
        chk_errors[inovice_filename].append(" |复核人名称非法| ")
    if payee and not re.findall(re.compile(r'[\u4e00-\u9fa5]+'),payee):
        chk_errors[inovice_filename].append(" |收款人名称非法| ")
    else:
        return True


def verify_expire(inovice_filename, invoicing_date):
    # 获取当前日期
    now = datetime.datetime.now()

    invoicing_date_str =  re.sub(r"[\u4e00-\u9fa5]+", r'/', invoicing_date,2).replace('日','')
    # 将日期字符串格式化为YYYY/MM/DD格式
    date = datetime.datetime.strptime(invoicing_date_str, "%Y/%m/%d")
    # 计算差额月数
    months_diff = (now.year - date.year) * 12 + (now.month - date.month)

    if months_diff > 3:
        chk_errors[inovice_filename].append(' |开票日期超过3个月| ')
        return False
    else:
        return True

def invoice_has_noerror(inovice_filename,invoicing_date, buytax_num,  selltax_num, buy_company_name,drawer,reviewer,payee):

    has_noerror = True

    if not verify_people_info(inovice_filename, drawer, reviewer, payee):
        has_noerror = False

    if not verify_expire(inovice_filename, invoicing_date):
        has_noerror = False

    if not verify_buycompany_ids(inovice_filename, buytax_num):
        has_noerror = False
        chk_errors[inovice_filename].append(' |公司纳税人识别号错误| ')
    if not verify_sellcompany_ids(inovice_filename, selltax_num):
        has_noerror = False
        chk_errors[inovice_filename].append(' |开票方纳税人识别号不合规| ')
    if not verify_company_name(inovice_filename, buy_company_name):
        has_noerror = False
        chk_errors[inovice_filename].append('|公司名称错误| ')
        return has_noerror
    else:
        return has_noerror

def print_perinvoice_chkres():
    if chk_errors:
        for pdffile in chk_errors:
            for error in chk_errors[pdffile]:
                print(f'!!!Warning :{pdffile} {error}')

def check_invoice_fee(oil_fee, catering_fee, comnicat_fee):
    if oil_fee < 500:
        print(f'!!!Warning : 交通总费用为{oil_fee},不满足月度报销500额度')
    if catering_fee < 680:
        print(f'!!!Warning : 餐饮总费用为{catering_fee},不满足月度报销680额度')
    if comnicat_fee < 300:
        print(f'!!!Warning : 通信总费用为{comnicat_fee},不满足月度报销300额度')

# GUI 界面
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from tkinter import scrolledtext
from tkinter import messagebox

def get_pdf(dir_path):
    # 定义文件名全局变量
    global files
    pdf_file = []
    for root, dirs, files in os.walk(dir_path):
        for name in files:
            if name.endswith('.pdf'):
                filepath = os.path.join(root, name)
                pdf_file.append(filepath)
    return pdf_file, files

def choose_folder():
    # 创建一个文件夹选择器
    global folder_path, pdfiles
    folder_path = filedialog.askdirectory()
    entry_folder_path.delete(0, tk.END)
    entry_folder_path.insert(0,folder_path)
    #var_folder_path.set(folder_path)

    pdfiles, totalfiles = get_pdf(folder_path)
    entry_pdf_info.delete(0, tk.END)
    var_pdf_info.set(f'一共发现{len(totalfiles)}个文件, 其中{len(pdfiles)}个PDF文件')

class myStdout():  # 重定向类
    def __init__(self):
        # 将其备份
        self.stdoutbak = sys.stdout
        self.stderrbak = sys.stderr
        # 重定向
        sys.stdout = self
        sys.stderr = self

    def write(self, info):
        # info信息即标准输出sys.stdout和sys.stderr接收到的输出信息
        # 在多行文本控件最后一行插入print信息
        scrolltext.insert('end', info)
        # 更新显示的文本，不加这句插入的信息无法显示
        scrolltext.update()
        # 始终显示最后一行，不加这句，当文本溢出控件最后一行时，不会自动显示最后一行
        scrolltext.see(tk.END)

    def restoreStd(self):
        # 恢复标准输出
        sys.stdout = self.stdoutbak
        sys.stderr = self.stderrbak

def read_pdf(folder_path):

    # 修改为自己的文件目录
    row = 1
    coin_sum = 0
    oil_cost_sum = 0
    catering_cost_sum = 0
    comucate_cost_sum = 0

    # 进度条初始值
    progress_bar['value'] = 0
    # 进度条最大值
    progress_bar['maximum'] = len(pdfiles)

    for pdffile in pdfiles:
        # 进度条更新
        progress_bar['value'] += 1
        # 画面更新
        root.update()
        print(pdffile)
        with pdfplumber.open(pdffile) as pdf:
            first_page = pdf.pages[0]
            pdf_text = first_page.extract_text()
            if '发票' not in pdf_text:
                continue
            print('--------------------------------------------------------')

            fapiaodaima = re_text(re.compile(r'发票代码(.*\d+)'), pdf_text)
            fapiaohaoma = re_text(re.compile(r'发票号码(.*\d+)'), pdf_text)
            kaipiaoriqi = re_text(re.compile(r'开票日期(.*)'), pdf_text)
            nashuishibie = re_text(re.compile(r'纳税人识别号\s*[:：]\s*([a-zA-Z0-9])\s*([a-zA-Z0-9]+)'), pdf_text)
            jiaoyan = re_text(re.compile(r'校\s*验\s*码\s*:([a-zA-Z0-9 ]+)'), pdf_text)
            fee = re.sub(r'小写.*[¥￥]','',re_text(re.compile(r'小写.*(.*[0-9.]+)'), pdf_text))
            shoukuanren = re_text(re.compile(r'收\s*款\s*人[:：]\s*[a-zA-Z0-9\u4e00-\u9fa5]+'), pdf_text)
            fuhe = re_text(re.compile(r'复\s*核[:：]\s*[a-zA-Z0-9\u4e00-\u9fa5]+'), pdf_text)
            kaipiaoren = re_text(re.compile(r'开\s*票\s*人[:：]\s*[a-zA-Z0-9\u4e00-\u9fa5]+'), pdf_text)
            buy_gongsi = re_text(re.compile(r'名\s*称\s*[:：]\s*([\u4e00-\u9fa5]+)'), pdf_text)

            invoice_code = fapiaodaima[5:] if fapiaodaima is not None else ""
            invoice_number = fapiaohaoma[5:] if fapiaohaoma is not None else ""
            invoicing_date = kaipiaoriqi[5:] if kaipiaoriqi is not None else ""
            purchasing_company = buy_gongsi[3:] if buy_gongsi is not None else ""
            payee = shoukuanren[4:] if shoukuanren is not None else ""
            reviewer = fuhe[3:] if fuhe is not None else ""
            drawer = kaipiaoren[4:] if kaipiaoren is not None else ""
            buyer_company_ids = nashuishibie[7:] if nashuishibie is not None else ""

            print(fapiaodaima)
            print(fapiaohaoma)
            print(kaipiaoriqi)
            print(buy_gongsi)
            print(nashuishibie)
            print(shoukuanren)
            print(fuhe)
            print(kaipiaoren)
            print(f'金额:{fee}')

            company = re.findall(re.compile(r'名.*称\s*[:：]\s*([\u4e00-\u9fa5]+)'), pdf_text)
            tax_num = re.findall(re.compile(r'纳税人识别号\s*[:：]\s*([a-zA-Z0-9])\s*([a-zA-Z0-9]+)'), pdf_text)
            tax_num = list(map(lambda eles: ''.join([ ele for ele in eles]), tax_num))

            is_oil = re.findall(re.compile(r'汽.*油'), pdf_text)
            is_catering = re.findall(re.compile(r'餐.*饮'), pdf_text)
            is_comucate = re.findall(re.compile(r'通.*信'), pdf_text)

            # 统计交通费
            if is_oil:
                oil_cost_sum += float(fee)

            # 统计餐饮费
            if is_catering:
                catering_cost_sum += float(fee)

            # 统计通信费
            if is_comucate:
                comucate_cost_sum += float(fee)

            if company:
                sell_gongsi = re_block(company[len(company) - 1])
                print(f'卖方公司：{sell_gongsi}')

            if tax_num:
                sell_taxnum = re_block(tax_num[len(tax_num) - 1])
                print(f'卖方公司纳税人识别号：{sell_taxnum}')

            checkinfo = '发票信息正确'
            if not invoice_has_noerror(pdffile, invoicing_date, nashuishibie,  sell_taxnum, buy_gongsi,drawer,reviewer,payee):
                checkinfo = chk_errors[pdffile]

            try:
                lst = [pdffile, invoice_code, invoice_number, invoicing_date, sell_gongsi, sell_taxnum, payee, reviewer, drawer,purchasing_company,buyer_company_ids, fee, checkinfo]
            except TypeError:
                print(">>>>>!!!info : 发票部分信息为空<<<<<")
            finally:
                # 填写信息入表
                for i in range(len(lst)):
                    sh.write(row, i, lst[i])
                row += 1

            # 累计金额
            coin_sum += float(fee)
            print('--------------------------------------------------------')

    print('#----------------------检查结果--------------------------#')
    print(f'交通费合计：{oil_cost_sum}')
    print(f'餐饮费合计：{catering_cost_sum}')
    print(f'通信费合计：{comucate_cost_sum}')
    print(f'合计金额：{coin_sum}')
    check_invoice_fee(oil_cost_sum, catering_cost_sum, comucate_cost_sum)
    print_perinvoice_chkres()
    sh.write(row, len(info_titles) - 1, f"合计：{coin_sum}")
    print('#-------------------------------------------------------#')

def export_to_excel():
    current_path = os.path.dirname(os.path.realpath(sys.argv[0]))
    savexlsx = '发票信息.xls'
    xlspath = '\\'.join((current_path, savexlsx))
    print(f"xls存放路径为{xlspath}")
    if os.path.exists(xlspath):
        os.remove(xlspath)
    tk.messagebox.showinfo(title='提示', message=f'{xlspath}导出成功！')
    # 保存
    wb.save('发票信息.xls')

root = tk.Tk()

# 设置文本框的标题为 "电子发票识别助手"
root.title("电子发票识别助手___Developed By ==> cheng Version:20231026")

root.geometry("1000x400")

# 创建一个标签,用于显示文件夹名称
tk.Label(root, text='选择路径: ', font= 20).place(x=50, y=25)
# 输入框显示文件夹路径
var_folder_path = tk.StringVar() #大模型生成
entry_folder_path = tk.Entry(root, textvariable=var_folder_path, fg='grey',font='15')
entry_folder_path.place(x=150, y=25, width=500, height=30 )
entry_folder_path.insert(0,'请选择文件存储路径：')

tk.Label(root, text='遍历情况: ', font = 20).place(x=50, y=75)
# 输入框显示PDF文件数量
var_pdf_info = tk.StringVar()
entry_pdf_info = tk.Entry(root, textvariable=var_pdf_info, fg='grey', font='15')
entry_pdf_info.place(x=150, y=75, width=500, height=30 )
entry_pdf_info.insert(0,'一共发现0个文件, 其中0个PDF文件')

tk.Label(root, text='进度情况: ', font = 20).place(x=50, y=125)
# 输入框显示PDF识别进度
var_pdf_progress = tk.StringVar()
# 进度条
progress_bar = ttk.Progressbar(root)
progress_bar.place(x=150, y=125, width=500, height=30)

# 创建一个按钮,用于选择文件夹
btn_selectfolder= tk.Button(root, text='浏览遍历文件夹', font = 18, command=choose_folder)
btn_selectfolder.place(x=700, y=23)
btn_selectfolder.config(width='15')

# 创建一个按钮,用于显示PDF文件信息
btn_verifypdf= tk.Button(root, text='识别PDF文件', font = 18, command=lambda:read_pdf(folder_path))
btn_verifypdf.place(x=700, y=73)
btn_verifypdf.config(width='15')

# 创建一个按钮,用于导出excel
btn_exportexcel= tk.Button(root, text='导出excel', font = 18, command=export_to_excel)
btn_exportexcel.place(x=700, y=123)
btn_exportexcel.config(width='15')

# 创建滚动文本框，来显示打印信息
scrolltext = scrolledtext.ScrolledText(root, width=100, height=15, font=('黑体', 10))
scrolltext.place(x=150,y=175)

# 创建一个 Text 对象
my_text = tk.Text(root)

# 设置文本框的文本内容
my_text.delete(1.0, tk.END)
my_text.insert(tk.END, "请确认被检测PDF文件均已被解密!!!")

# 创建一个 Text 对象并将其添加到应用程序中
my_text.place(x=150,y=380)
my_text.config(font=("Courier", 10), fg="red", bg='light grey',width=30)

# 实例化重定向类
mystd = myStdout()

root.mainloop()

# 恢复标准输出
mystd.restoreStd()
