#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Author  : Tai cu Shi ba
# Copyright (C) Tai cu Shi ba, All Rights Reserved
# @File    : pdf_electronic_invoice_recognition.py
# @IDE     : PyCharm
import os
import re
import pdfplumber


def re_search_text(bt, text):
    """使用search 匹配text"""
    m1 = re.search(bt, text)
    if m1 is not None:
        return m1.group(1)


def re_finditer_text(compile_bt, text):
    """查询所有符合规则的信息"""
    return list(compile_bt.finditer(text))


def read_pdf_invoice(pdf_file_path):
    """获取发票信息"""
    invoice_info_dict = []
    # 读取单个文件，只读取pdf格式的发票
    if os.path.splitext(pdf_file_path)[1] == '.pdf':
        with (pdfplumber.open(pdf_file_path) as pdf):
            first_page = pdf.pages[0]
            pdf_text = first_page.extract_text()
            if '发票' in pdf_text:
                invoice_code = re_search_text(r'发票代码[:：]\s*(.*\d+)', pdf_text)
                invoice_num = re_search_text(r'发票号码[:：]\s*(.*\d+)', pdf_text)
                invoice_date = re_search_text(r'开票日期[:：]\s*(.*)', pdf_text)
                # 名称
                both_name_rule = re.compile(r'名\s*称\s*[:：]\s*([\u4e00-\u9fa5]+)')
                both_name_set = re_finditer_text(both_name_rule, pdf_text)
                # 识别号
                tax_code_rule = re.compile(r'纳税人识别号\s*[:：]\s*([a-zA-Z0-9]\s*[a-zA-Z0-9]+)')
                tax_code_set = re_finditer_text(tax_code_rule, pdf_text)
                # 项目名称
                project_name_rule = re.compile(r'\*(.*?)\*([\u4e00-\u9fa5]+.*?)')
                project_name_set = re_finditer_text(project_name_rule, pdf_text)
                # 总金额
                total_fee = re_search_text(r'（小写）¥(\d+\.\d+)', pdf_text)
                # 数据符合要求
                if all([invoice_num, invoice_date, project_name_set, total_fee]) and len(
                        both_name_set) == 2 and len(tax_code_set) == 2:
                    # 税方信息
                    buyer = both_name_set[0].group(1)
                    seller = both_name_set[1].group(1)
                    # 唯一税号
                    buyer_tax_code = tax_code_set[0].group(1)
                    seller_tax_code = tax_code_set[1].group(1)
                    # 项目名称
                    total_project_name = '、'.join(set([i.group() for i in project_name_set]))
                    invoice_info_dict = {"invoice_code": invoice_code,
                                         "invoice_num": invoice_num, "invoice_date": invoice_date,
                                         "buyer": buyer, "buyer_tax_code": buyer_tax_code,
                                         "seller": seller, "seller_tax_code": seller_tax_code,
                                         "total_project_name": total_project_name,
                                         "total_fee": total_fee}
                    print('--------------------------------------------------------')
                    print(f"发票代码:{invoice_code}\n"
                          f"发票号码:{invoice_num}\n开票日期:{invoice_date}\n"
                          f"购买公司:{buyer}\n购买公司税号:{buyer_tax_code}\n"
                          f"销售公司:{seller}\n销售公司税号:{seller_tax_code}\n"
                          f"项目名称:{total_project_name}\n费用总计:{total_fee}")
                    print('--------------------------------------------------------')
    return invoice_info_dict


if __name__ == '__main__':
    path = os.getcwd()
    path += f'\\test_file\\test2.pdf'
    res = read_pdf_invoice(path)
