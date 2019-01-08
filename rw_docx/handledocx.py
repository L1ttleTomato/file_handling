"""
    从word文档读取表格中特定内容及相关内容
"""

import os
from docx import Document


def get_doc_path_list(num = 0):
    """ 获取当前目录的doc绝对路径列表 """
    path = os.getcwd()
    doc_path= path + '\doc_file'
    doc_list = []
    for file in os.listdir(doc_path):
        if file[-4:] == 'docx':
            doc_list.append(file)
    return doc_path + '\\' + doc_list[num]


def get_table(file_path):
    """获取doc文件的table列表"""
    document = Document(file_path)
    tables = document.tables
    return tables


def get_doc_tables():
    return get_table(get_doc_path_list())


def get_param_value(param_name, offset, tables=get_doc_tables()):
    """
    :param param_name: 参数名
    :param offset: 偏移量
    :param tables: 默认的文档表格tables
    :return: 键值对的字典
    """
    age_list = {}
    for table in tables:
        row = len(table.rows)
        col = len(table.columns)
        for i in range(row):
            for j in range(col):
                name = table.cell(i, j).text     # 获取i行j列的cell的内容
                if name == param_name:
                    age_list[name] = table.cell(i, j+offset).text
    return age_list





if __name__ == '__main__':
    # tables = get_doc_tables()
    # age_list = {}                  # 存放搜索到的值的字典

    # 遍历文档中所有表格的cell
    # for table in tables:
    #     row = len(table.rows)
    #     col = len(table.columns)
    #     for i in range(row):
    #         for j in range(col):
    #             name = table.cell(i, j).text     # 获取i行j列的cell的内容
    #             if name in ['张三', '张五']:
    #                 age_list[name] = table.cell(i, j+3).text
    # print(age_list)
    #

    print(get_param_value('张五', 2))
    #
    # cell01 = tables[0].cell(0, 0).text
    # cell02 = tables[0].columns[0].cells
    # # for cell in cell02:
    # #     print(cell.text)
    # print(help(tables[0]))






"""
    Help on Table in module docx.table object:

class Table(docx.shared.Parented)
 |  Proxy class for a WordprocessingML ``<w:tbl>`` element.
 |  
 |  Method resolution order:
 |      Table
 |      docx.shared.Parented
 |      builtins.object
 |  
 |  Methods defined here:
 |  
 |  __init__(self, tbl, parent)
 |      Initialize self.  See help(type(self)) for accurate signature.
 |  
 |  add_column(self, width)
 |      Return a |_Column| object of *width*, newly added rightmost to the
 |      table.
 |  
 |  add_row(self)
 |      Return a |_Row| instance, newly added bottom-most to the table.
 |  
 |  cell(self, row_idx, col_idx)
 |      Return |_Cell| instance correponding to table cell at *row_idx*,
 |      *col_idx* intersection, where (0, 0) is the top, left-most cell.
 |  
 |  column_cells(self, column_idx)
 |      Sequence of cells in the column at *column_idx* in this table.
 |  
 |  row_cells(self, row_idx)
 |      Sequence of cells in the row at *row_idx* in this table.
 |  
 |  ----------------------------------------------------------------------
 |  Data descriptors defined here:
 |  
 |  alignment
 |      Read/write. A member of :ref:`WdRowAlignment` or None, specifying the
 |      positioning of this table between the page margins. |None| if no
 |      setting is specified, causing the effective value to be inherited
 |      from the style hierarchy.
 |  
 |  autofit
 |      |True| if column widths can be automatically adjusted to improve the
 |      fit of cell contents. |False| if table layout is fixed. Column widths
 |      are adjusted in either case if total column width exceeds page width.
 |      Read/write boolean.
 |  
 |  columns
 |      |_Columns| instance representing the sequence of columns in this
 |      table.
 |  
 |  rows
 |      |_Rows| instance containing the sequence of rows in this table.
 |  
 |  style
 |      Read/write. A |_TableStyle| object representing the style applied to
 |      this table. The default table style for the document (often `Normal
 |      Table`) is returned if the table has no directly-applied style.
 |      Assigning |None| to this property removes any directly-applied table
 |      style causing it to inherit the default table style of the document.
 |      Note that the style name of a table style differs slightly from that
 |      displayed in the user interface; a hyphen, if it appears, must be
 |      removed. For example, `Light Shading - Accent 1` becomes `Light
 |      Shading Accent 1`.
 |  
 |  table
 |      Provide child objects with reference to the |Table| object they
 |      belong to, without them having to know their direct parent is
 |      a |Table| object. This is the terminus of a series of `parent._table`
 |      calls from an arbitrary child through its ancestors.
 |  
 |  table_direction
 |      A member of :ref:`WdTableDirection` indicating the direction in which
 |      the table cells are ordered, e.g. `WD_TABLE_DIRECTION.LTR`. |None|
 |      indicates the value is inherited from the style hierarchy.
 |  
 |  ----------------------------------------------------------------------
 |  Data descriptors inherited from docx.shared.Parented:
 |  
 |  __dict__
 |      dictionary for instance variables (if defined)
 |  
 |  __weakref__
 |      list of weak references to the object (if defined)
 |  
 |  part
 |      The package part containing this object

None

"""