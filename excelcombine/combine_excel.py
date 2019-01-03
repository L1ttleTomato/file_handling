import pandas as pd
import os


def get_xls_list():
    """ 获取当前目录的xls文件名列表 """
    path = os.path.abspath('.')
    all_excel_name = []
    for file in os.listdir(path):
        if file[-3:] == 'xls':
            all_excel_name.append(file)
    return all_excel_name


"""
    read_excel()参数介绍
    io ：excel 路径
    sheetname：默认是sheetname为0，返回多表使用sheetname=[0,1]，若sheetname=None是返回全表 。注意：int/string返回的是dataframe，而none和list返回的是dict of dataframe。
    header ：指定作为列名的行，默认0，即取第一行，数据为列名行以下的数据；若数据不含列名，则设定 header = None；
    skiprows：省略指定行数的数据
    skip_footer：省略从尾部数的行数据
    index_col ：指定列为索引列，也可以使用 u’string
    names：指定列的名字，传入一个list数据

"""


def combine_xls(xlslist):
    """ 将xls文件合并，参数为xls文件名的列表"""
    xls_list = xlslist
    combine = []
    for xls in xls_list:
        data = pd.read_excel(xls)
        combine.append(data)
        writer = pd.ExcelWriter('combined.xls')
        pd.concat(combine).to_excel(writer, sheet_name='sheet1', index=False)
    writer.save()


def main():
    combine_xls(get_xls_list())


if __name__ == '__main__':
    main()
