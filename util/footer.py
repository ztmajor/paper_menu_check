# encoding: utf-8
"""
@author: xie yufeng
@software: PyCharm
@file: footer.py
@time: 2021/4/23 19:41
"""
import os
import fitz
import docx

from util import *


def find_page_number(pdf_file, keyword):
    page_list = []
    with fitz.open(pdf_file) as doc:
        for idx, page in enumerate(doc, start=1):
            if keyword in page.getText():
                # print(f"page {idx} include keyword {keyword}!!")
                page_list.append(idx)
    return page_list


def check_alignment(docx_file):
    check_msg = []
    roma_nums_uppercase = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X']
    document = docx.Document(docx_file)
    for section in document.sections:
        for parg in section.footer.paragraphs:
            if parg.text:
                if parg.text in roma_nums_uppercase:
                    if str(parg.alignment) != 'CENTER (1)':
                        # print('页码 {} 未居中'.format(parg.text))
                        check_msg.append('页码 {} 未居中'.format(parg.text))
    return check_msg


def check_footer_nums(pdf_file, start_page, end_page):
    roma_nums_uppercase = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X']
    check_msg = []
    with fitz.open(pdf_file) as doc:
        for idx, page in enumerate(doc, start=1):
                if start_page <= idx < end_page:
                    cur_footer = roma_nums_uppercase[idx-start_page]
                    if '图目录' in page.getText():
                        cur_sec = '图目录'
                    elif '表目录' in page.getText():
                        cur_sec = '表目录'
                    elif '目录' in page.getText():
                        cur_sec = '目录'
                    if cur_footer in page.getText():
                        # print("{} 页码：{}".format(cur_sec, cur_footer))
                        check_msg.append("{} 页码：{}".format(cur_sec, cur_footer))
                    else:
                        # print("{} 缺少页码：{} 或者页码格式非大写罗马字符".format(cur_sec, cur_footer))
                        check_msg.append("{} 缺少页码：{} 或者页码格式非大写罗马字符".format(cur_sec, cur_footer))
    return check_msg


def check_footer(pdf, docx):
    '''
    不检查是否存在目录、表目录、图目录，默认都存在
    :param pdf: pdf file path
    :param docx: docx file path
    :return: None
    '''
    stand_page_list = find_page_number(pdf, '目录')
    p_list = find_page_number(pdf, '第 1 章')
    footer_num_msg = check_footer_nums(pdf, min(stand_page_list), min(p_list))
    alignment_msg = check_alignment(docx)
    return footer_num_msg + alignment_msg


if __name__ == '__main__':
    doc_path = os.getcwd() + '/../paper/硕士学位论文正文_1.doc'
    pdf_path = doc_path.replace('doc', 'pdf')
    docx_path = doc_path.replace('doc', 'docx')
    msg = check_footer(pdf_path, docx_path)
    print(msg)
