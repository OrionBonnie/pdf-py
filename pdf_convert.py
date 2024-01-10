try:
    import subprocess
    from pathlib import Path
    from time import sleep
    from pdf2word import Converter
    from re import match
    from os import cpu_count, listdir
    from os.path import splitext
    from msvcrt import getch
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
    print("模块载入完毕")
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
    print(f"开始转化 {pdf}")
    sleep(1)
    in_pdf_path = pdf
    out_docx_path = Path(r".\out")
    converter = Converter(in_pdf_path)
    cpu_num = cpu_count()
    if not out_docx_path.exists():
        out_docx_path.mkdir()
    if splitext(in_pdf_path)[-1] == ".pdf":
        docx_filename = str(in_pdf_path.replace("pdf", "docx"))
        docx_filename = docx_filename.replace("in", "out")
        converter.convert(docx_filename, multiprocessing=cpu_num)
    elif splitext(in_pdf_path)[-1] == ".PDF":
        docx_filename = str(in_pdf_path.replace("PDF", "docx"))
        docx_filename = docx_filename.replace("in", "out")
        converter.convert(docx_filename, multiprocessing=cpu_num)
    converter.close()


def main():
    for pdf_file in get_pdf_list():
        convert2docx(pdf_file)
        print(f"转换完成，共有 {len(get_pdf_list())} 个文件被转换成 docx， 格式即背景颜色可能会有小毛病")
        print("比如边距会被改变，这个可以在布局一栏中选择边距然后选择正常即可")
        print("还有就是表格的位置可能会出现问题，单击表格然后在表格布局菜单中选择自适应，选择自适应当前窗口")
        print("按任意键退出")
        getch()
