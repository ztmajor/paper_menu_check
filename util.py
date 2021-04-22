# encoding: utf-8
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
                # print(f"page {idx} include keyword {keyword}!!")
                page_list.append(idx)
    return page_list


def check_alignment(docx_file):
    roma_nums_uppercase = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X']
    document = docx.Document(docx_file)
    for section in document.sections:
        for parg in section.footer.paragraphs:
            if parg.text:
                if parg.text in roma_nums_uppercase:
                    if str(parg.alignment) != 'CENTER (1)':
                        print('页码 {} 未居中'.format(parg.text))


def check_footer_nums(pdf_file, start_page, end_page):
    roma_nums_uppercase = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X']
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
                        print("{} 页码：{}".format(cur_sec, cur_footer))
                    else:
                        print("{} 缺少页码：{} 或者页码格式非大写罗马字符".format(cur_sec, cur_footer))


def check_footer(pdf, docx):
    '''
    不检查是否存在目录、表目录、图目录，默认都存在
    :param pdf: pdf file path
    :param docx: docx file path
    :return: None
    '''
    stand_page_list = find_page_number(pdf, '目录')
    p_list = find_page_number(pdf, '第1章')
    check_footer_nums(pdf, min(stand_page_list), min(p_list))
    check_alignment(docx)


if __name__ == '__main__':
    # 这里只能使用完整绝对地址，相对地址找不到文件，且，只能用“\\”，不能用“/”，哪怕加了 r 也不行，涉及到将反斜杠看成转义字符。
    doc_path = r'W:\ZJU\课程\春学期\写作指导\作业\hw\paper\硕士学位论文正文_1.doc'
    docx_path = r'C:\\Users\\xieyu\\Desktop\\jupyter project\\论文课程作业\\硕士学位论文正文_1.docx'
    pdf_path = r'C:\\Users\\xieyu\\Desktop\\jupyter project\\论文课程作业\\硕士学位论文正文_1.pdf'
    # doc2docx(doc_path, docx_path)
    # for p in docx_file.paragraphs:
    #     print(p.text)
    # doc2pdf(doc_path, pdf_path)
    stand_page_list = find_page_number(pdf_path, '目录')
    p_list = find_page_number(pdf_path, '第1章')
    # print('final page:', max(p_list) - min(stand_page_list) + 1)
    check_footer(pdf_path, docx_path)