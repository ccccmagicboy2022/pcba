#!d:\cccc2020\TOOL\python-3.9.1-embed-amd64\python.exe
BOM_CSV_FILE = "BOM_PCB_AIRFLOW_METER_B_2022-06-19.csv"
BOM_OUTPUT_FILE = "BOM_PCB_AIRFLOW_METER_B_2022-06-19.xlsx"

import os
import xlsxwriter

part_list = []
part_list_ordered = []
sum_lcsc = 0.0

bom_file_path = os.path.join(os.getcwd(), ".", BOM_CSV_FILE,)
print("Open the BOM file: {0:s}".format(bom_file_path))
with open(bom_file_path, "r", encoding="utf-16-le") as ff:
    for line in ff.readlines():
        if "ID\tName\tDesignator\tFootprint\tQuantity" not in line:
            #print(line)
            part = line.strip("\n").replace('"', "").split("\t")
            num_needed = int(part[4])
            if ("" != part[9]):
                single_price = float(part[9])
            else:
                print(f'{part[8]} miss price')
            total = single_price * num_needed
            sum_lcsc += total
            del part[0] ##ID
            part_list.append(part)

print(sum_lcsc)
print(f"Size of part list is {len(part_list)}")
# 按Designator排序
part_list.sort(key=lambda part_list: part_list[1])  ##Designator

i = 0
# 输出最终的列表
for part in part_list:
    i += 1
    part = [i,] + part
    #print(part)
    part_list_ordered.append(part)

print(part_list_ordered)

# 准备输出XLSX文件
bom_file_path2 = os.path.join(os.getcwd(), ".", BOM_OUTPUT_FILE,)
if os.path.exists(bom_file_path2):
    os.remove(bom_file_path2)
    print('del the output file\r\n')

workbook = xlsxwriter.Workbook(bom_file_path2)
# 创建主工作表
worksheet = workbook.add_worksheet("BOM")
worksheet.repeat_rows(0)  # 第一行重复出现
worksheet.protect()  # 保护本表

# 以下是标题的格式
header_format = workbook.add_format(
    {"border": 1, "text_wrap": True, "locked": True, "font_size": 13, "bold": True,}
)
header_format.set_align("left")
header_format.set_align("vcenter")
# 以下是普通单元格的格式
cell_format = workbook.add_format(
    {"border": 1, "text_wrap": True, "locked": True, "font_size": 12,}
)
cell_format.set_align("left")
cell_format.set_align("vcenter")

# 写bom.xlsx
# 先写标题
worksheet.write_row(
    0,
    0,
    [
        "ID",
        "Name",
        "Designator",
        "Footprint",
        "Quantity",
        "Manufacturer Part",
        "Manufacturer",
        "Supplier",
        "Supplier Part",
        #"Price",
    ],
    header_format,
)
# 设置列宽
worksheet.set_column("A:A", 4)
worksheet.set_column("B:B", 24)
worksheet.set_column("C:C", 33)
worksheet.set_column("D:D", 33)
worksheet.set_column("E:E", 12)
worksheet.set_column("F:F", 23)
worksheet.set_column("G:G", 17)
worksheet.set_column("H:H", 12)
worksheet.set_column("I:I", 13)
worksheet.set_column("J:J", 12)
# 设置标题行高
worksheet.set_row(0, 40)
# 设置打印的格式
# https://xlsxwriter.readthedocs.io/page_setup.html
worksheet.set_landscape()
worksheet.set_paper(9)  # A4
worksheet.fit_to_pages(1, 0)
worksheet.set_margins(left=0.3, right=0.3, top=0.75, bottom=0.75)

# 设置页眉
worksheet.set_header(
    f"&L&C&15{BOM_CSV_FILE}&R",
)
# 设置页脚
worksheet.set_footer("&CPage &P of &N")

i = 0
for i in range(len(part_list_ordered)):
    xx_str = ""
    worksheet.set_row(i + 1, 40)  # 设置行高

    ##worksheet.write_row(i + 1, 0, part_list_ordered[i][:10], cell_format)
    worksheet.write_row(i + 1, 0, part_list_ordered[i][:9], cell_format)
    worksheet.write_url(
        "I{0:d}".format(i + 2),
        "https://so.szlcsc.com/global.html?k={0:s}".format(
            part_list_ordered[i][8]
        ),
        cell_format,
        string="{0:s}".format(part_list_ordered[i][8]),
    )

# 关闭文件
workbook.close()
