import os
import subprocess
import tqdm
from time import sleep
from msvcrt import getch
from pdf_pdf2word import get_pdf_list
from pdf_table2excel import choose_page

print("基础模块已导入")
try:
    import pdfplumber
except ModuleNotFoundError:
    list_mod_all = subprocess.run("pip list --disable-pip-version-check", capture_output=True)
    list_mod_all = list_mod_all.stdout.decode("utf-8")
    if list_mod_all.find("pdfplumber") == -1:
        subprocess.run("pip install pdfplumber")
        import pdfplumber
finally:
    print("PDF I/O流模块已导入")


def pdf_extract_img(pdf_path, resolution):
    selected, pdf_password = choose_page(pdf_path)
    if not selected:
        return

    with pdfplumber.open(pdf_path, password=pdf_password) as pdf_io:
        print("开始...共提取 {} 页，请耐心等待".format(len(selected)))
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        for page in tqdm.tqdm(selected):
            image_page = pdf_io.pages[page-1]
            image = image_page.to_image(resolution=int(resolution))
            save_path = os.path.join(r".\out", f"{base_name}-第{page}页.png")
            image.save(save_path)


def main():
    if not os.path.exists(r".\out"):
        os.mkdir(r".\out")
    else:
        for file in os.listdir(r".\out"):
            os.remove(os.path.join(r".\out", file))
            
    res = input("请输入提取图片的分辨率(100-1000, 分辨率越大，图片越清晰，但提取速度会降低)[回车默认250]: ")
    if not res.isdigit():
        res = 250
    for pdf_path in get_pdf_list():
        pdf_extract_img(pdf_path, res)
        
    total_img = len(os.listdir(r".\out"))
    print("已完成，共提取了 {} 张图片".format(total_img))
    sleep(1.5)
    os.system(r"start .\out")
    print("按任意键退出关闭此窗口")
    getch()

if __name__ == "__main__":
    main()