import os
import subprocess
import sys
from pathlib import Path
from time import sleep
from re import match
from os import cpu_count, listdir
from os.path import splitext
from msvcrt import getch
from tqdm import tqdm
try:
    from pdf2word import Converter
except ModuleNotFoundError:
    command = "pip list"
    list_result = subprocess.run(command, capture_output=True)
    list_result = list_result.stdout.decode("utf-8")
    if list_result.find("pdf2word") == -1:
        print("未找到 与 pdf 转换 word 的相关模块,开始下载 pdf2word 模块")
        subprocess.run("pip install pdf2word")
        from pdf2word import Converter
    else:
        print("程序检查模块出现异常，终止脚本")
        print("按任意键退出本脚本")
        getch()
        exit()
finally:
    print("pdf转word模块已导入")
    sleep(2)


def get_pdf_list():
    path = r".\in"
    pdf_regex = r".*(PDF|pdf)$"
    in_path = Path(path)
    if not in_path.exists():
        print()
        print("输入目录不存在，请在本脚本所在的目录下创建一个 in 文件夹，并将PDF文件放入 in 内")
        print("按任意键退出本脚本")
        getch()
        exit()
    elif len(listdir(in_path)) == 0:
        print()
        print("输入目录内未找到任何文件")
        print("按任意键退出本脚本")
        getch()
        exit()
    to_convert_list = list()
    for file in in_path.iterdir():
        file = str(file)
        if match(pdf_regex, file):
            to_convert_list.append(file)
    if len(to_convert_list) == 0:
        print()
        print("输入目录内没有找到任何 PDF 文件，此脚本只可用于转换PDF文件")
        print("按任意键退出本脚本")
        getch()
        exit()
    print(f"共发现 {len(to_convert_list)}个 PDF 文件")
    return to_convert_list


def convert2docx(pdf):
    sleep(1)
    in_pdf_path = pdf
    converter = Converter(in_pdf_path)
    cpu_num = cpu_count()
    docx_path = ""
    if splitext(in_pdf_path)[-1] == ".pdf":
        docx_filename = str(in_pdf_path.replace("pdf", "docx"))
        docx_path = docx_filename.replace("in", "out")
    elif splitext(in_pdf_path)[-1] == ".PDF":
        docx_filename = str(in_pdf_path.replace("PDF", "docx"))
        docx_path = docx_filename.replace("in", "out")
    converter.convert(docx_path, multiprocessing=cpu_num)
    converter.close()


def main():
    if not os.path.exists(r".\out"):
        os.mkdir(r".\out")
    else:
        for file in os.listdir(r".\out"):
            os.remove(r".\out\\" + file)
    for pdf_file in get_pdf_list():
        convert2docx(pdf_file)
    print("转换完成，共转换 {} 个docx, 格式及背景颜色可能会有小毛病".format(os.listdir(r'.\out')))
    sleep(1.5)
    os.system(r"start .\out")
    print("按任意键退出关闭此窗口")
    getch()
