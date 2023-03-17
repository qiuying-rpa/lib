import os
import re
import win32com.client

from qiuying.models.excel import ExcelObj


def open_excel(file_path, sheet_name=1, password=None, visible=True):
    """
    打开已有 Excel
    :param file_path:
    :param sheet_name:

    :param password:
    :param visible:
    :return:
    """
    # todo excel_type  office/wps
    if not os.path.exists(file_path):
        raise Exception('目标文件不存在，请确认后重试')
    xlapp = None
    workbook = None
    if "EXCEL.EXE" in os.popen('tasklist /fi "IMAGENAME eq EXCEL.EXE"').read():
        xlapp = win32com.client.GetActiveObject('Excel.Application')
        names = [i.Name for i in xlapp.Workbooks]
        path, name = os.path.split(file_path)
        if file_path in names:
            workbook = xlapp.Workbooks(file_path)
        elif name in names:
            workbook = xlapp.Workbooks(name)
        else:
            if password:
                workbook = xlapp.Workbooks.Open(file_path, UpdateLinks=False, ReadOnly=False, Format=None,
                                                Password=password)
            else:
                workbook = xlapp.Workbooks.Open(file_path)
    if not workbook:
        xlapp = win32com.client.DispatchEx('Excel.Application')
        if password:
            workbook = xlapp.Workbooks.Open(file_path, UpdateLinks=False, ReadOnly=False, Format=None,
                                            Password=password)
        else:
            workbook = xlapp.Workbooks.Open(file_path)

    xlapp.Visible = visible
    xlapp.DisplayAlerts = False
    workbook.Activate()
    sheet_name_list = [i.Name for i in workbook.Worksheets]
    if isinstance(sheet_name, str):
        if sheet_name not in sheet_name_list:
            raise Exception("无法找到该名称的工作表，请核对工作表名称")
    else:
        if sheet_name > len(sheet_name_list):
            raise Exception("工作表序号超过实际工作表个数，请核对工作表名称")
    worksheet = workbook.Sheets(sheet_name)
    worksheet.Activate()
    return ExcelObj(xlapp, workbook, worksheet)


def create_excel(file_path, visible=True):
    """
    新建 Excel
    :param file_path:
    :param visible:
    :return:
    """
    if os.path.exists(file_path):
        print("目标文件已存在，无需新建！")
    try:
        xlapp = win32com.client.GetActiveObject('Excel.Application')
    except Exception as e:
        xlapp = win32com.client.Dispatch('Excel.Application')
    xlapp.Visible = visible
    xlapp.DisplayAlerts = False
    workbook = xlapp.Workbooks.Add()
    workbook.SaveAs(file_path)
    worksheet = workbook.Sheets(1)
    return ExcelObj(xlapp, workbook, worksheet)


def get_worksheet_name(excel_obj):
    """
    获取所有sheet名称
    :param excel_obj:
    :return:
    """
    return [i.Name for i in excel_obj.workbook.Worksheets]


def switch_worksheet(excel_obj, sheet_name):
    """
    切换工作表
    :param excel_obj:
    :param sheet_name:
    :return:
    """
    sheet_name_list = get_worksheet_name(excel_obj)
    if isinstance(sheet_name, str):
        if sheet_name not in sheet_name_list:
            raise Exception("无法找到该名称的工作表，请核对工作表名称")
    else:
        if sheet_name > len(sheet_name_list):
            raise Exception("工作表序号超过实际工作表个数，请核对工作表名称")
    excel_obj.worksheet = excel_obj.workbook.Sheets(sheet_name)
    excel_obj.worksheet.Activate()


def create_worksheet(excel_obj, sheet_name):
    """
    新建 sheet 页
    :param excel_obj:
    :param sheet_name:
    :return:
    """
    sheet_name_list = get_worksheet_name(excel_obj)
    if sheet_name in sheet_name_list:
        raise Exception("新建工作表名称与当前存在的工作表名称相同！")
    worksheet = excel_obj.workbook.Worksheets.Add()
    worksheet.Name = sheet_name
    excel_obj.worksheet = worksheet


def delete_worksheet(excel_obj, sheet_name):
    """
    删除 sheet 页
    :param excel_obj:
    :param sheet_name:
    :return:
    """
    sheet_name_list = get_worksheet_name(excel_obj)
    if isinstance(sheet_name, str):
        if sheet_name not in sheet_name_list:
            raise Exception("无法找到该名称的工作表，请核对工作表名称")
    else:
        if sheet_name > len(sheet_name_list):
            raise Exception("工作表序号超过实际工作表个数，请核对工作表名称")
    excel_obj.workbook.Worksheets(sheet_name).Delete()


def save(excel_obj, file_path=None):
    """
    保存/另存为 Excel 文件
    :param excel_obj:
    :param file_path:
    :return:
    """
    # file_format_map = {
    #     'xlsx': 51,
    #     'xls': 56,
    #     'csv': 6,
    #     'html': 44,
    #     'xml': 51,
    #     'txt': -4158,
    #     'xlsm': 52,
    # }
    if file_path is None:
        excel_obj.workbook.Save()
    else:
        # file_format = file_path.split('.')[-1]
        # file_format_code = file_format_map.get(file_format, None)
        # wb_file_format = excel_obj.workbook.Name.split(".")[-1]   # 当前workbook后缀
        # file_format = wb_file_format if file_format_code is None else file_format_code
        # file_format_code = file_format_map.get(file_format, None)

        # 对于现有文件，默认采用上一次指定的文件格式；对于新文件，默认采用当前所用 Excel 版本的格式。
        excel_obj.workbook.SaveAs(file_path)


def get_value(excel_obj, region_text, attr="真实值"):
    """
    读取Excel内容
    :param excel_obj:
    :param region_text:
    :param attr:[“真实值”, "显示值", "公式"]
    :return:
    """
    region = _get_region(excel_obj, region_text)
    if attr == "显示值":
        data = tuple(map(lambda row: tuple(map(lambda col: col.Text, row.Columns)), region.Rows))
    elif attr == "真实值":
        data = region.Value
    else:
        data = region.Formula
    if not isinstance(data, tuple):
        return data
    if len(data) == 1:
        data = data[0]
        if len(data) == 1:
            return data[0]
        return list(data)
    else:
        return [i[0] if len(i) == 1 else list(i) for i in data]


def set_value(excel_obj, region_text, value, str_flag=False):
    """
    写入内容至单元格区域
    :param excel_obj:
    :param region_text: 起始单元格
    :param value:
    :param str_flag:
    :return:
    """
    cell = _get_region(excel_obj, region_text).Cells(1, 1)
    col = cell.Column
    row = cell.Row

    if isinstance(value, (list, tuple)):
        type_check = all([isinstance(i, (list, tuple)) for i in value])
        if not type_check:
            raise Exception("列表(元组)格式错误，仅支持二维列表(元组)，请检查。")
        cols = max(len(i) for i in value)
        value = list(map(lambda x: x + [""] * (cols - len(x)), value))
        rows = len(value)
        write_region_text = "{0},{1}:{2},{3}".format(row, col, row + rows - 1, col + cols - 1)
    elif isinstance(value, (str, float, int)):
        write_region_text = region_text
    else:
        raise Exception("不支持的数据类型写入。")
    region = _get_region(excel_obj, write_region_text)
    if str_flag:
        region.NumberFormat = '@'
        region.Value = value
    else:
        region.Value = value


def close_excel(excel_obj, close_type="关闭指定Excel", save_flag=True, file_path=None):
    """
    关闭Excel
    :param close_type: ["关闭指定Excel", "关闭所有Excel"]
    :param excel_obj:
    :param save_flag:
    :param file_path:
    :return:
    """
    if close_type == "关闭指定Excel":
        if excel_obj is None:
            raise Exception("需要给定excel对象")
        if save_flag:
            save(excel_obj, file_path)
        excel_obj.workbook.Close()
        if not excel_obj.xlapp.Workbooks.Count:
            excel_obj.xlapp.Quit()
    else:
        if excel_obj:
            excel_obj.xlapp.Quit()
        else:
            os.system('taskkill /f /im EXCEL.exe')


def get_row_count(excel_obj):
    """
    读取Excel总行数
    :param excel_obj:
    :return:
    """
    return _get_region(excel_obj, "all").Rows.Count


def get_column_count(excel_obj):
    """
    读取Excel总列数
    :param excel_obj:
    :return:
    """
    return _get_region(excel_obj, "all").Columns.Count


def insert_rows(excel_obj, row, row_num=1):
    """
    插入行
    :param excel_obj:
    :param row: 指定行前插入一行
    :param row_num:
    :return:
    """
    region = _get_region(excel_obj, "{0}:{0}".format(row), used=False)
    for i in range(row_num):
        region.Insert()


def insert_columns(excel_obj, column, column_num=1):
    """
    插入列
    :param excel_obj:
    :param column:指定列前插入一行
    :param column_num:
    :return:
    """
    if str(column).isdigit():
        column = int(column)
    if isinstance(column, int):
        column = tran_col_location(column)
    region = _get_region(excel_obj, "{0}:{0}".format(column), used=False)
    for i in range(column_num):
        region.Insert()


def delete_rows(excel_obj, start_row, end_row):
    """
    删除行
    :param excel_obj:
    :param start_row:
    :param end_row:
    :return:
    """
    region = _get_region(excel_obj, "{0}:{1}".format(start_row, end_row), used=False)
    region.Delete()


def delete_columns(excel_obj, start_col, end_row):
    """
    删除列
    :param excel_obj:
    :param start_col:
    :param end_row:
    :return:
    """
    if str(start_col).isdigit():
        start_col = int(start_col)
    if isinstance(start_col, int):
        start_col = tran_col_location(start_col)

    if str(end_row).isdigit():
        end_row = int(end_row)
    if isinstance(end_row, int):
        end_row = tran_col_location(end_row)

    region = _get_region(excel_obj, "{0}:{1}".format(start_col, end_row), used=False)
    region.Delete()


def set_region_shape(excel_obj, region_text, width=-1, height=-1):
    """
    区域长宽设置
    :param excel_obj:
    :param region_text:
    :param width:-1为不设置， AUTO为自适应
    :param height:-1为不设置， AUTO为自适应
    :return:
    """
    region = _get_region(excel_obj, region_text)
    if str(width).upper() == "AUTO":
        region.Columns.AutoFit()

    if str(height).upper() == "AUTO":
        region.Rows.AutoFit()

    if width > 0:
        region.ColumnWidth = width
    if height > 0:
        region.RowHeight = height


def set_region_format(excel_obj, region_text, format_text=None):
    # TODO 其他格式待开发
    """
    单元格格式设置
    :param excel_obj:
    :param region_text:
    :param format_text:
    :return:
    """
    region = _get_region(excel_obj, region_text)
    # 仅设置单元格数字格式
    num_format_map = {
        "文本": "@",
        "年-月-日": "yyyy-m-d",
        "两位小数": "0.00",
        "百分比": "0.00%"
    }
    if format_text:
        region.NumberFormatLocal = num_format_map.get(format_text, format_text)


def merge_cells(excel_obj, region_text):
    """
    合并单元格
    :param excel_obj:
    :param region_text:
    :return:
    """
    region = _get_region(excel_obj, region_text)
    region.Merge()


def unmerge_cells(excel_obj, region_text):
    """
    拆分单元格
    :param excel_obj:
    :param region_text:
    :return:
    """
    region = _get_region(excel_obj, region_text)
    region.UnMerge()


def excel_sort(excel_obj, sort_args):
    """
    排序
    :param excel_obj:
    :param sort_args:
    :return:
    """
    """sort attr
    MatchCase - 设置为 True 可执行区分大小写的排序，或设置为 False 以执行不区分大小写的排序
    Header - 指定第一行是否包含标题信息。
        xlNo      2   默认值
        xlYes     1
        xlGuess   0   如果希望 Excel 确定标题，可以指定  
    Orientation - 指定排序方向 
        xlSortColumns  1      按列  
        xlSortRows     2      按行    默认值
    SortFields -  该对象代表与 Sort 对象关联的排序字段的集合
    SortMethod - 指定中文排序方法。 
        xlPinYin	1	按字符的汉语拼音顺序排序。 这是默认值。 
        xlStroke	2	按每个字符的笔划数排序。
    SetRange - 设置排序发生区域。（自定义）
        1  扩展选定区域    默认值
        2  以当前区域排序  
    """
    """fields attr
    Key  - 指定排序字段，该字段确定要排序的值。 range对象
    SortOn - 设置要排序的单元格的属性 
        SortOnValues         0    值
        SortOnCellColor      1    单元格颜色
        SortOnFontColor      2    字体颜色
        SortOnIcon           3    图标
    Order - 确定关键字所指定的值的排序次序
        xlAscending     1     升序
        xlDescending    2     降序
    DataOption - 指定如何对 SortField 对象中指定的范围中的文本进行排序
        xlSortNormal            0   分别对数字和文本数据进行排序。  默认值
        xlSortTextAsNumbers     1   将文本作为数字型数据进行排序。
    """
    # sort_args = {
    #     "SortFields": [
    #         {
    #             "region": "A",
    #             "SortOn": 0,
    #             "Order": 1,
    #             "DataOption": 0
    #         },
    #         {
    #             "region": "B",
    #             "SortOn": 0,
    #             "Order": 2,
    #             "DataOption": 0
    #         }
    #     ],
    #     "SortAttr": {
    #         "SetRange": "",
    #         "MatchCase": False,
    #         "Header": 2,
    #         "Orientation": 1,
    #         "SortMethod": 1
    #     }
    # }
    excel_obj.worksheet.Sort.SortFields.Clear()
    for sort_field in sort_args["SortFields"]:
        region = _get_region(excel_obj, sort_field["region"])
        excel_obj.worksheet.Sort.SortFields.Add(Key=region,
                                                SortOn=sort_field["SortOn"],
                                                Order=sort_field["Order"],
                                                DataOption=sort_field["DataOption"])
    # if len(sort_args["SortFields"]) > 1:
    #     excel_obj.worksheet.Sort.SetRange(_get_region(excel_obj, "all"))
    # else:
    #     excel_obj.worksheet.Sort.SetRange(region)
    excel_obj.worksheet.Sort.SetRange(_get_region(excel_obj, "all"))  # 默认扩展选定区域
    excel_obj.worksheet.Sort.MatchCase = sort_args["SortAttr"]["MatchCase"]
    excel_obj.worksheet.Sort.Header = sort_args["SortAttr"]["Header"]
    excel_obj.worksheet.Sort.Orientation = sort_args["SortAttr"]["Orientation"]
    excel_obj.worksheet.Sort.SortMethod = sort_args["SortAttr"]["SortMethod"]
    excel_obj.worksheet.Sort.Apply()


def filter_col(excel_obj, region_text, criteria1, option="筛选值", criteria2=None):
    # TODO 其他筛选项
    """
    筛选列数据
    :param excel_obj:
    :param region_text:
    :param criteria1:
    :param option: 仅支持[筛选值, 条件1和条件2的逻辑与, 条件1和条件2的逻辑或]
    :param criteria2:
    :return:
    """
    cell = _get_region(excel_obj, region_text)
    cell.Select()
    if excel_obj.worksheet.AutoFilterMode:
        region = excel_obj.worksheet.AutoFilter.Range
    else:
        region = excel_obj.xlapp.Selection.CurrentRegion

    col_start = region.Cells(1, 1).Column
    col_target = cell.Column

    option_map = {
        "条件1和条件2的逻辑与": 1,  # xlAnd
        "条件1和条件2的逻辑或": 2,  # xlOr
        "显示最高值项（条件1中指定的项数）": 3,  # xlTop10Items
        "显示最高值项（条件1中指定的百分数）": 5,  # xlTop10Percent
        "显示最低值项（条件1中指定的项数）": 4,  # xlBottom10Items
        "显示最低值项（条件1中指定的百分数）": 6,  # xlBottom10Percent
        "筛选值": 7,  # xlFilterValues
        "单元格颜色": 8,  # xlFilterCellColor
        "字体颜色": 9,  # xlFilterFontColor
        "筛选图标": 10,  # xlFilterIcon
        "动态筛选": 11,  # xlFilterDynamic
    }
    if option != "筛选值":
        region.AutoFilter(Field=col_target - col_start + 1, Criteria1=criteria1, Operator=option_map[option],
                          Criteria2=criteria2)
    else:
        region.AutoFilter(Field=col_target - col_start + 1, Criteria1=criteria1)


def remove_filter(excel_obj):
    """
    清除筛选
    :param excel_obj:
    :return:
    """
    excel_obj.worksheet.ShowAllData()


def auto_fill(excel_obj, source, destination, fill_type="默认"):
    """
    自动填充
    :param excel_obj:
    :param source:
    :param destination:
    :param fill_type:
    :return:
    """
    region = _get_region(excel_obj, source, used=False)
    if ":" not in source:
        start = source
    else:
        start = source.split(":")[0]
    if ":" not in destination:
        end = destination
    else:
        end = destination.split(":")[1]
    dst = _get_region(excel_obj, f"{start}:{end}", used=False)

    xl_fill_map = {
        "默认": 0,
        "复制值和格式": 1,
        "复制格式": 3,
        "复制值": 4,
        "乘法计算": 9,
        "加法计算": 10,
    }
    xl_fill = xl_fill_map.get(fill_type, fill_type)
    if xl_fill not in list(range(11)):
        raise Exception("填充类型请参照说明")
    region.AutoFill(Destination=dst, Type=xl_fill)
    """
    xlFillCopy 1 将源区域的值和格式复制到目标区域，如有必要可重复执行。
    xlFillDays 5 将星期中每天的名称从源区域扩展到目标区域中。格式从源区域复制到目标区域，如有必要可重复执行。
    xlFillDefault 0 Excel 确定用于填充目标区域的值和格式。
    xlFillFormats 3 只将源区域的格式复制到目标区域，如有必要可重复执行。
    xlFillMonths 7 将月名称从源区域扩展到目标区域中。格式从源区域复制到目标区域，如有必要可重复执行。
    xlFillSeries 2 将源区域中的值扩展到目标区域中，形式为系列（如，“1, 2”扩展为“3, 4, 5”）。格式从源区域复制到目标区域，如有必要可重复执行。
    xlFillValues 4 只将源区域的值复制到目标区域，如有必要可重复执行。
    xlFillWeekdays 6 将工作周每天的名称从源区域扩展到目标区域中。格式从源区域复制到目标区域，如有必要可重复执行。
    xlFillYears 8 将年从源区域扩展到目标区域中。格式从源区域复制到目标区域，如有必要可重复执行。
    xlGrowthTrend 10 将数值从源区域扩展到目标区域中，假定源区域的数字之间是乘法关系（如，“1, 2,”扩展为“4, 8, 16”，假定每个数字都是前一个数字乘以某个值的结果）。格式从源区域复制到目标区域，如有必要可重复执行。
    xlLinearTrend 9 将数值从源区域扩展到目标区域中，假定数字之间是加法关系（如，“1, 2,”扩展为“3, 4, 5”，假定每个数字都是前一个数字加上某个值的结果）。格式从源区域复制到目标区域，如有必要可重复执行。
    """


def print_set(excel_obj, orientation='纵向', paper_size=None, wide=False, tall=False):
    """
    打印设置
    :param excel_obj:
    :param orientation:
    :param paper_size:
    :param wide: 是否缩放为一页
    :param tall: 是否缩放为一页
    :return:
    """
    paper_size_map = {
        "A4": 9,  # xlPaperA4
        "A3": 8,  # xlPaperA3
        "A5": 11,  # xlPaperA5
        "B4": 12,  # xlPaperB4
        "B5": 13,  # xlPaperB5
    }
    orientation = 1 if orientation == '纵向' else 2
    excel_obj.worksheet.PageSetup.Orientation = orientation
    if paper_size:
        excel_obj.worksheet.PageSetup.PaperSize = paper_size_map.get(paper_size)
    if wide or tall:
        excel_obj.worksheet.PageSetup.Zoom = False
    if wide:
        excel_obj.worksheet.PageSetup.FitToPagesWide = wide
    if tall:
        excel_obj.worksheet.PageSetup.FitToPagesTall = tall


def copy_region(excel_obj, region_text):
    """
    复制区域
    :param excel_obj:
    :param region_text:
    :return:
    """
    region = _get_region(excel_obj, region_text)
    region.Copy()


def paste_region(excel_obj, region_text, paste_option="粘贴全部内容"):
    """
    粘贴内容到单元格区域
    :param excel_obj:
    :param region_text:
    :param paste_option:
    :return:
    """
    region = _get_region(excel_obj, region_text)
    paste_option_map = {
        "粘贴全部内容": -4104,  # xlPasteAll
        "粘贴除边框外的全部内容": 7,  # xlPasteAllExceptBorders
        "将粘贴所有内容，并且将合并条件格式": 14,  # xlPasteAllMergingConditionalFormats
        "使用源主题粘贴全部内容": 13,  # xlPasteAllUsingSourceTheme
        "粘贴复制的列宽": 8,  # xlPasteColumnWidths
        "粘贴批注": -4144,  # xlPasteComments
        "粘贴复制的源格式": -4122,  # xlPasteFormats
        "粘贴公式": -4123,  # xlPasteFormulas
        "粘贴公式和数字格式": 11,  # xlPasteFormulasAndNumberFormats
        "粘贴有效性": 6,  # xlPasteValidation
        "粘贴值": -4163,  # xlPasteValues
        "粘贴值和数字格式": 12,  # xlPasteValuesAndNumberFormats
    }
    region.PasteSpecial(Paste=paste_option_map[paste_option])


def find_cell_address(excel_obj, text, region_text="all", look_in="值", look_at="全部匹配", match_case=False):
    """
    查找内容所在单元格位置
    :param excel_obj:
    :param text:
    :param region_text:
    :param look_in:
    :param look_at:
    :param match_case: 如果为 True，则搜索区分大小写。 默认值为 False。
    :return:
    """
    look_in_map = {
        "批注": -4144,  # xlComments
        "公式": -4123,  # xlFormulas
        "值": -4163,  # xlValues
    }
    look_at_map = {
        "部分匹配": 2,  # xlPart
        "全部匹配": 1,  # xlWhole
    }
    region = _get_region(excel_obj, region_text)
    match_cell = region.Find(What=text, After=list(region)[-1], LookIn=look_in_map[look_in],
                             LookAt=look_at_map[look_at], MatchCase=match_case)
    match_cell_address = []
    while match_cell:
        if match_cell.Address.replace("$", "") in match_cell_address:
            break
        match_cell_address.append(match_cell.Address.replace("$", ""))
        match_cell = region.FindNext(match_cell)
    return match_cell_address


def tran_col_location(col_location):
    """
    列数字与字母转换
    :param col_location:
    :return:
    """
    alphabets = [
        'A', 'B', 'C', 'D', 'E', 'F', 'G',
        'H', 'I', 'J', 'K', 'L', 'M', 'N',
        'O', 'P', 'Q', 'R', 'S', 'T',
        'U', 'V', 'W', 'X', 'Y', 'Z',
    ]
    if isinstance(col_location, str) and col_location.isalpha():
        col_str = col_location.upper()
        col_length = len(col_str)
        col_int = sum(
            [(alphabets.index(_) + 1) * (26 ** (col_length - _index - 1)) for _index, _ in enumerate(col_str)])
        return col_int
    elif isinstance(col_location, int) and col_location > 0:
        col_int = col_location
        col_str = ""
        while col_int > 26:
            col_int, rem = divmod(col_int, 26)
            col_str = alphabets[rem - 1] + col_str
        col_str = alphabets[col_int - 1] + col_str
        return col_str
    else:
        raise Exception("不符合列名转换规则")


def _get_region(excel_obj, region_text, used=True):
    """
    通过文本获取Range or Rows or Columns
    :param excel_obj:
    :param region_text:
    :param used:
    :return:
    """
    ws = excel_obj.worksheet
    region_text = region_text.replace("：", ":").replace("，", ",").upper()
    used_region = ws.UsedRange
    left = used_region.Column
    top = used_region.Row
    right = left + used_region.Columns.Count - 1
    bottom = top + used_region.Rows.Count - 1
    if region_text == "ALL":
        return ws.Range(ws.Cells(1, 1), ws.Cells(bottom, right))
    elif region_text == "SELECTED":
        return excel_obj.xlapp.Selection

    # 正则
    # 单元格 '[A-Z]+\d+$'       A2, B3, C7 ...
    # 行列索引单元格  '(\d+),(\d+)$'   row,col
    # 行 '\d+$',               2, 3, 4...
    # 列 '[A-Z]+$'          A, C, F...

    if ":" in region_text:
        start, end = region_text.split(":")
        if all([re.match(r'[A-Z]+\d+$', i) for i in [start, end]]):
            return ws.Range(region_text)
        elif all([re.match(r'(\d+),(\d+)$', i) for i in [start, end]]):
            row1, col1 = [int(i) for i in start.split(",")]
            row2, col2 = [int(i) for i in end.split(",")]
            return ws.Range(ws.Cells(row1, col1), ws.Cells(row2, col2))
        elif all([re.match(r'\d+$', i) for i in [start, end]]):
            if used:
                return ws.Range(ws.Cells(int(start), 1), ws.Cells(int(end), right))
            else:
                return ws.Rows(region_text)
        elif all([re.match('[A-Z]+$', i) for i in [start, end]]):
            if used:
                return ws.Range(f'{start}{1}:{end}{bottom}')
            else:
                return ws.Columns(region_text)
        else:
            raise Exception(f"{region_text}无法识别的单元格区域")
    else:
        if re.match(r'[A-Z]+\d+$', region_text):
            return ws.Range(region_text)
        elif re.match(r'(\d+),(\d+)$', region_text):
            row, col = [int(i) for i in region_text.split(',')]
            return ws.Cells(row, col)
        elif re.match(r'\d+$', region_text):
            if used:
                return ws.Range(ws.Cells(int(region_text), 1), ws.Cells(int(region_text), right))
            else:
                return ws.Rows(int(region_text))
        elif re.match('[A-Z]+$', region_text):
            if used:
                return ws.Range(f'{region_text}{1}:{region_text}{bottom}')
            else:
                return ws.Columns(int(region_text))
        else:
            raise Exception(f"{region_text}无法识别的单元格区域")


if __name__ == '__main__':
    pass
