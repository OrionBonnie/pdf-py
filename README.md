# pdf-py
简单的Python小脚本，用于批量转换 PDF 至 Word、提取 PDF 中的表格至 Excel 以及 PDF 图片提取
A simple Python script for batch converting PDFs to Word, extracting tables from PDFs to Excel, and extracting images from PDFs.

所需环境： Python 3.10及以上 和 JDK 1.8
Requirements: Python 3.10 or higher, and JDK 1.8.

在开始之前，请先把所有的 PDF 文件都放入 in 文件夹里， 如果没有 in 文件夹，先在脚本所在的路径创建。若没有 in 目录，脚本会报错
Before you begin, please place all PDF files into the `in` folder. If the `in` folder does not exist, create it in the same directory where the script is located. The script will throw an error if the `in` directory is missing.

之后就可以运行 "运行我.bat" 这个文件，根据提示走
Afterward, you can run the "RunMe.bat" file and follow the prompts.

最后所有的转化后的文件都会输出到脚本目录中一个名为 out 的文件夹中
All converted files will be output to a folder named `out` within the script directory.



---
本脚本仍未完善，但可正常工作。转换后的文件会出现格式小毛病，但都可手动修改。
如果出现检查模块下载时卡住，如果有代理的可以代理一下pip，具体如何代理可以百度搜一下

This script is not yet fully polished, though it functions correctly. The converted files may contain minor formatting glitches, but these can all be corrected manually.
If the script gets stuck while downloading verification modules, users with a proxy server can configure pip to use that proxy; you can search on Google for specific instructions on how to do this.
