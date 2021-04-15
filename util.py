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


def doc2docx(doc_path, docx_path):
    # doc文件另存为docx
    doc = win32com.client.Dispatch('kwps.Application')
    doc_file = doc.Documents.Open(doc_path)
    doc_file.SaveAs(docx_path, 12, False, "", True, "", False, False, False, False)  # 转换后的文件,12代表转换后为docx文件
    # doc.SaveAs(r"F:\\***\\***\\appendDoc\\***.docx", 12) # 或直接简写
    # 注意SaveAs会打开保存后的文件，有时可能看不到，但后台一定是打开的
    doc_file.Close()
    doc.Quit()


def read_docx(docx_path):
    return docx.Document(docx_path)


def doc2pdf(doc_file, pdf_file):
    wdFormatPDF = 17
    word = Dispatch('kwps.Application')
    doc = word.Documents.Open(doc_file)
    doc.SaveAs(pdf_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()


def find_page_number(pdf_file, keyword):
    page_list = []
    with fitz.open(pdf_file) as doc:
        for idx, page in enumerate(doc, start=1):
            if keyword in page.getText():
                print(f"page {idx} include keyword {keyword}!!")
                page_list.append(idx)
    return page_list


if __name__ == '__main__':
    # 这里只能使用完整绝对地址，相对地址找不到文件，且，只能用“\\”，不能用“/”，哪怕加了 r 也不行，涉及到将反斜杠看成转义字符。
    doc_path = r'W:\hw\paper\硕士学位论文正文_1.doc'
    docx_path = r'W:\hw\paper\硕士学位论文正文_1.docx'
    pdf_path = r'W:\hw\paper\硕士学位论文正文_1.pdf'
    # doc2docx(doc_path, docx_path)
    # for p in docx_file.paragraphs:
    #     print(p.text)
    # doc2pdf(doc_path, pdf_path)
    stand_page_list = find_page_number(pdf_path, '第 1 章')
    p_list = find_page_number(pdf_path, '图 2.1')
    print('final page:', max(p_list) - max(stand_page_list) + 1)
