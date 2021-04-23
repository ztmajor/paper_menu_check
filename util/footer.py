# encoding: utf-8
"""
@author: duke mei
@contact: mylovezju@163.com
@software: PyCharm
@file: util.py
@time: 2021/4/15 8:55
"""
from os import remove
from os.path import splitext
import win32com.client
from win32com.client import Dispatch, constants
import docx
import fitz
import pydoc_data


def find_page_number(pdf_file, keyword):
    page_list = []
    with fitz.open(pdf_file) as doc:
        for idx, page in enumerate(doc, start=1):
            if keyword in page.getText():
                # print(f"page {idx} include keyword {keyword}!!")
                page_list.append(idx)
    return page_list


def check_alignment(docx_file):
    all_good = True
    document = docx.Document(docx_file)
    for section in document.sections:
        for parg in section.footer.paragraphs:
            if '目录' in section.header.paragraphs[0].text:
                if parg.text:
                    if str(parg.alignment) != 'CENTER (1)':
                        print('页码 {} 未居中'.format(parg.text))
                        all_good = False
    if all_good:
        print('目录页码已全部居中')


def check_footer_nums(pdf_file, start_page, end_page):
    roma_nums_uppercase = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X']
    all_good = True
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
                    data = [item.strip('\n') for item in page.getText().split(' ')]
                    if cur_footer in data:
                        print("{} 页码：{}".format(cur_sec, cur_footer))
                    else:
                        print("{} 缺少页码：{} 或者页码格式非大写罗马字符".format(cur_sec, cur_footer))
                        all_good = False
    return all_good


def check_footer(pdf, docx):
    '''
    不检查是否存在目录、表目录、图目录，默认都存在
    :param pdf: pdf file path
    :param docx: docx file path
    :return: None
    '''
    stand_page_list = find_page_number(pdf, '目录')
    p_list = find_page_number(pdf, '第1章')
    if check_footer_nums(pdf, min(stand_page_list), min(p_list)):
        check_alignment(docx)
    else:
        print('请修改错误的页码')






if __name__ == '__main__':
    # 这里只能使用完整绝对地址，相对地址找不到文件，且，只能用“\\”，不能用“/”，哪怕加了 r 也不行，涉及到将反斜杠看成转义字符。
    doc_path = r'C:\\Users\\xieyu\\Desktop\\jupyter project\\论文课程作业\\test2.doc'
    docx_path = r'C:\\Users\\xieyu\\Desktop\\jupyter project\\论文课程作业\\test2.docx'
    pdf_path = r'C:\\Users\\xieyu\\Desktop\\jupyter project\\论文课程作业\\test2.pdf'
    check_footer(pdf_path, docx_path)


