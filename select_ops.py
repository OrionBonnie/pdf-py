from msvcrt import getch
print('''
这里是 PDF 各类转换操作的选择菜单，请选择你想要的操作
1. PDF 批量转换 Word
2. PDF 表格提取至 Excel
3. PDF 图片提取
## 注意: 如果 out 文件夹内仍存在文件，将会在选择操作被清空，若存在重要文件，请及时移出 ##
''')
choose = input("请选择对应的操作数字: ")
if choose == "1":
    import pdf_convert
    pdf_convert.main()
elif choose == "2":
    import pdf_table2excel
    pdf_table2excel.main()
elif choose == "3":
    import pdf_imagout
    pdf_imagout.main()
else:
    print("未知操作, 按任意键退出")
    getch()
