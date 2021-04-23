# encoding: utf-8
import win32com.client
from win32com.client import Dispatch, constants
import docx
import pythoncom


def get_word_com():
    try:
        doc_com = Dispatch('kwps.Application')
    except pythoncom.com_error:
        doc_com = Dispatch('Word.Application')

    return doc_com


def doc2docx(doc_path, docx_path):
    doc = get_word_com()
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
    # word = Dispatch('kwps.Application')
    word = get_word_com()
    doc = word.Documents.Open(doc_file)
    doc.SaveAs(pdf_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()
