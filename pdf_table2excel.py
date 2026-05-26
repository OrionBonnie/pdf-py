import os
import subprocess
import csv
from msvcrt import getch
from pdf_pdf2word import get_pdf_list
from tqdm import tqdm
from time import sleep

print("基本模块已导入")
try:
    check_java = subprocess.run("java -version", capture_output=True)
except FileNotFoundError:
    print("请检查是否安装好 JDK 1.8 且配置好环境变量")
    print("按任意键退出")
    getch()
    exit()
finally:
    print("依赖检查完成, 为可用状态")

try:
    import pypdf
except ModuleNotFoundError:
    list_mod_all = subprocess.run("pip list --disable-pip-version-check", capture_output=True)
    list_mod_all = list_mod_all.stdout.decode("utf-8")
    if list_mod_all.find("pypdf") == -1:
        subprocess.run("pip install pypdf")
        import pypdf
finally:
    print("pdf读取模块已导入")

try:
    import tabula
except ModuleNotFoundError:
    list_mod_all = subprocess.run("pip list --disable-pip-version-check", capture_output=True)
    list_mod_all = list_mod_all.stdout.decode("utf-8")
    if list_mod_all.find("tabula-py") == -1:
        subprocess.run("pip install tabula-py jpype1")
        import tabula
finally:
    print("列表读取模块已导入")

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Border, Side
except ModuleNotFoundError:
    list_mod_all = subprocess.run("pip list --disable-pip-version-check", capture_output=True)
    list_mod_all = list_mod_all.stdout.decode("utf-8")
    if list_mod_all.find("openpyxl") == -1:
        subprocess.run("pip install openpyxl")
        from openpyxl import Workbook
finally:
    print("excel格式模块已导入")


def is_empty_row(in_row):
    result_list = list()
    for test_null in in_row:
        result_list.append(not bool(test_null))
    return all(result_list)


def choose_page(pdf_path):
    reader = pypdf.PdfReader(pdf_path)
    pdf_password = None
    if reader.is_encrypted:
        print(f"\n发现加密文件: {pdf_path}")
        while True:
            pwd = input("请输入该PDF的密码 (直接按回车可跳过该文件): ")
            if not pwd:
                print("已跳过此加密文件。")
                return None, None
            if reader.decrypt(pwd):
                print("密码正确，准备读取页数...")
                pdf_password = pwd
                break
            else:
                print("密码错误，请重新输入！")
    
    pdf_pages_total = len(reader.pages)
    print('''
        当前正在对 {} 操作(共 {} 页):
        1. 全部提取
        2. 指定页数范围(默认从第1页开始, 指定范围时用空格把两个页数隔开, 如: 3 6)
        3. 指定具体页数(用空格把多个页数隔开, 如: 1 3 6 7)
        '''.format(os.path.basename(pdf_path), pdf_pages_total))
        
    choice = input("请输入指定的选项数字(1,2,3): ")
    selected_pages = list()
    
    if choice == "1":
        for i in range(pdf_pages_total):
            selected_pages.append(i+1)
    elif choice == "2":
        page_between = input("请输入你想要提取的范围: ")
        split = page_between.split()
        if len(split) != 2:
            print("你只可输入两个数字作为范围")
            print("按任意键退出")
            getch()
            exit()
        for i in range(int(split[0]), int(split[1]) + 1):
            selected_pages.append(i)
    elif choice == "3":
        page_between = input("请输入你想要提取的页码: ")
        split = page_between.split()
        # 修正了之前的bug，保证转换出的都是整数，防止比对失败
        selected_pages = [int(p) for p in split]
        
    if choice in ["2", "3"]:
        if max(selected_pages) > pdf_pages_total:
            print("页数超过PDF文件的最大页数, 您输入了大于PDF页数的页码")
            print("按任意键退出")
            getch()
            exit()
            
    return selected_pages, pdf_password


def pdf_table2csv(pdf_path):
    selected, pdf_password = choose_page(pdf_path)
    if not selected:
        return None

    if not os.path.exists(r".\csv"):
        os.mkdir(r".\csv")
        
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    csv_file = f"{base_name}.csv"
    csv_path = os.path.join(r".\csv", csv_file)
    
    kwargs = {"output_format": "csv", "pages": selected, "lattice": True}
    if pdf_password:
        kwargs["password"] = pdf_password
        
    tabula.convert_into(pdf_path, csv_path, **kwargs)
    print(f"开始提取转换 {pdf_path}")
    return csv_path


def csv2xlsx(csv_file):
    rows_content = list()
    with open(csv_file, "r", encoding="utf-8") as csv_io:
        csv_reader = csv.reader(csv_io)
        for csv_row in csv_reader:
            if not is_empty_row(csv_row):
                rows_content.append(csv_row)

    web_book = Workbook()
    sheet = web_book.active
    row_num = 1
    column_max_num = 1
    border = Border(left=Side(style="thin"), right=Side(style="thin"), 
                    top=Side(style="thin"), bottom=Side(style="thin"))
    alignment = Alignment(horizontal="left", vertical="center", wrapText=False, shrinkToFit=True)

    for row in tqdm(rows_content):
        sheet.append(row)
        for i in range(len(row)):
            col = i + 1
            cell = sheet.cell(row=row_num, column=col)
            cell.border = border
            cell.alignment = alignment
            if col > column_max_num:
                column_max_num = col
        row_num += 1

    init_column_ascii = ord("A")
    for column in range(column_max_num):

        sheet.column_dimensions[chr(init_column_ascii)].width = 18
        init_column_ascii += 1

    base_name = os.path.splitext(os.path.basename(csv_file))[0]
    xlsx_save_path = os.path.join(r".\out", f"{base_name}.xlsx")
    web_book.save(xlsx_save_path)


def main():
    if not os.path.exists(r".\out"):
        os.mkdir(r".\out")
    else:
        for file in os.listdir(r".\out"):
            os.remove(os.path.join(r".\out", file))
            
    for pdf_file in get_pdf_list():
        csv_path = pdf_table2csv(pdf_file)
        if csv_path:
            csv2xlsx(csv_path)
            
    if os.path.exists(r".\csv"):
        for file in os.listdir(r".\csv"):
            os.remove(os.path.join(r".\csv", file))
        os.rmdir(r".\csv")

    print('\n转换已完成，共从 {} 个文件中提取了表格'.format(len(os.listdir(r".\out"))))
    print("脚本为了表格输出整洁默认开启字体适应缩放，若一个单元格字数过多，该单元格内的字体也会变小，双击单元格即可恢复原样")
    sleep(1.5)
    os.system(r"start .\out")
    print("按任意键退出关闭此窗口")
    getch()

if __name__ == "__main__":
    main()