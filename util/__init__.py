# encoding: utf-8
"""
@author: duke mei
@contact: mylovezju@163.com
@software: PyCharm
@file: __init__.py
@time: 2021/4/23 19:51
"""
from util.catalog import check_catalog
from util.footer import check_footer
from util.util import doc2pdf, doc2docx

__all__ = [
    'doc2pdf',
    'doc2docx',
    'check_catalog',
    'check_footer'
]
