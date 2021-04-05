# -*- coding: utf-8 -*-
# @Time    : 2021/4/5 16:11
# @Author  : carlosliu
# @File    : pdf_files_merge.py

"""
处理 PDF 文件合并的小工具
"""

import easygui
import sys
import codecs
from PyPDF2 import PdfFileReader, PdfFileMerger
import logging
import time

logging.basicConfig(level=logging.INFO,
                    filename='pdf_files_merge.log',
                    filemode='a',
                    format='%(asctime)s - %(pathname)s[line:%(lineno)d] - %(levelname)s: %(message)s')


def files_merge(file_list, output_file):
    try:
        pdf_merger = PdfFileMerger()
        for file in file_list:
            f = codecs.open(file, 'rb')
            pdf_reader = PdfFileReader(f)
            if pdf_reader.isEncrypted:
                continue
            pdf_merger.append(pdf_reader)
            f.close()

        pdf_merger.write(output_file)
        pdf_merger.close()
    except BaseException as e:
        logging.error('合并PDF文件错误', exc_info=True)
        easygui.exceptionbox()
        pass
    logging.error('合并PDF文件成功', exc_info=True)
    easygui.msgbox("PDF文件合并完成")


if __name__ == '__main__':
    files_dir_path = []
    while True:
        file_dir_path = easygui.fileopenbox(msg='选择要合并的PDF文件(多选)', title='浏览文件夹', default="*.pdf", filetypes=["\\*.pdf"], multiple=False)
        if file_dir_path is None:
            easygui.msgbox('未选择PDF文件', title='提示')
            sys.exit()
        files_dir_path.append(file_dir_path)
        reply = easygui.buttonbox(msg="是否还有其他文件需要继续合并", choices=["是", "否"])
        if reply == "否":
            break

    if files_dir_path is None:
        easygui.msgbox('未选择PDF文件', title='提示')
        sys.exit()
    else:
        easygui.msgbox("要处理的PDF文件共有 {} 个\n\n选择保存的文件路径".format(len(files_dir_path)), title='提示')

    logging.info("要处理的PDF文件共有 {} 个".format(len(files_dir_path)))
    output_dir_path = easygui.diropenbox(msg="选择保存的文件路径", title="浏览文件夹")

    if output_dir_path is None:
        easygui.msgbox('未选择文件路径', title='提示')
        sys.exit()
    output_file = output_dir_path + "//output_{}.pdf".format(time.time())

    files_merge(file_list=files_dir_path, output_file=output_file)
