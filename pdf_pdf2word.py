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

# 检查并导入 pdf 转 word 模块
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

# 检查并导入 PDF 密码处理模块
try:
    from pypdf import PdfReader
except ModuleNotFoundError:
    print("未找到处理 PDF 密码所需的 pypdf 模块，开始下载...")
    subprocess.run("pip install pypdf")
    from pypdf import PdfReader

finally:
    print("PDF相关模块已准备完毕")
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
    pdf_password = None
    
    #检查是否存在密码并验证
    try:
        reader = PdfReader(in_pdf_path)
        if reader.is_encrypted:
            print(f"\n发现加密文件: {in_pdf_path}")
            while True:
                pwd = input("请输入该PDF的密码 (直接按回车可跳过该文件): ")
                if not pwd:
                    print("已跳过此加密文件。")
                    return  # 直接跳过转换
                
                # 尝试用输入的密码解密
                if reader.decrypt(pwd):
                    print("密码正确，开始转换...")
                    pdf_password = pwd
                    break
                else:
                    print("密码错误，请重新输入！")
    except Exception as e:
        print(f"检查PDF密码时发生错误 ({in_pdf_path}): {e}")
        return

    if pdf_password:
        converter = Converter(in_pdf_path, password=pdf_password)
    else:
        converter = Converter(in_pdf_path)

    cpu_num = cpu_count()
    
    file_name = os.path.basename(in_pdf_path)
    file_name_without_ext = splitext(file_name)[0]
    docx_path = os.path.join(r".\out", file_name_without_ext + ".docx")

    converter.convert(docx_path, multiprocessing=cpu_num)
    converter.close()


def main():
    if not os.path.exists(r".\out"):
        os.mkdir(r".\out")
    else:
        for file in os.listdir(r".\out"):
            os.remove(os.path.join(r".\out", file))
            
    pdf_files = get_pdf_list()
    for pdf_file in pdf_files:
        convert2docx(pdf_file)
        
    print("\n转换完成，共生成 {} 个docx, 格式及背景颜色可能会有小毛病".format(len(os.listdir(r'.\out'))))
    sleep(1.5)
    os.system(r"start .\out")
    print("按任意键退出关闭此窗口")
    getch()

if __name__ == "__main__":
    main()