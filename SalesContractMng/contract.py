from xlrd import open_workbook as wb
import xlwt
import os.path


class Line:

    @staticmethod
    def title(value):
        print("---------------------------" + value + "---------------------------")


class SheetWizard:

    def __init__(self, url):
        try:
            self.workbook = wb(url)
            self.mainsheet = self.get_main_sheet()
        except Exception:
            return

    def get_main_sheet(self):
        return self.workbook.sheet_by_index(0)

    @staticmethod
    def get_rows_by_condition(sheet, col_name, operator, value):
        index = 0
        res_list = []
        # 找到对应列名的列
        if '大于' == operator:
            for i in range(sheet.ncols):
                if sheet.cell_value(0, i) == col_name:
                    index = i
                    break
            for i in range(1, sheet.nrows):
                if value != "" and sheet.cell_value(i, index) != "" and float(value) < sheet.cell_value(i, index):
                    res_list.append(i)
        if '大于等于' == operator:
            for i in range(sheet.ncols):
                if sheet.cell_value(0, i) == col_name:
                    index = i
                    break
            for i in range(1, sheet.nrows):
                if value != "" and sheet.cell_value(i, index) != "" and float(value) <= sheet.cell_value(i, index):
                    res_list.append(i)
        if '小于' == operator:
            for i in range(sheet.ncols):
                if sheet.cell_value(0, i) == col_name:
                    index = i
                    break
            for i in range(1, sheet.nrows):
                if value != "" and sheet.cell_value(i, index) != "" and float(value) > sheet.cell_value(i, index):
                    res_list.append(i)
        if '小于等于' == operator:
            for i in range(sheet.ncols):
                if sheet.cell_value(0, i) == col_name:
                    index = i
                    break
            for i in range(1, sheet.nrows):
                if value != "" and sheet.cell_value(i, index) != "" and float(value) >= sheet.cell_value(i, index):
                    res_list.append(i)
        if '等于' == operator:
            for i in range(sheet.ncols):
                if sheet.cell_value(0, i) == col_name:
                    index = i
                    break
            for i in range(sheet.nrows):
                if value != "" and sheet.cell_value(i, index) != "" and float(value) == sheet.cell_value(i, index):
                    res_list.append(i)
        if '不等于' == operator:
            for i in range(sheet.ncols):
                if sheet.cell_value(0, i) == col_name:
                    index = i
                    break
            for i in range(sheet.nrows):
                if value != "" and sheet.cell_value(i, index) != "" and float(value) != sheet.cell_value(i, index):
                    res_list.append(i)
        if '是' == operator:
            for i in range(sheet.ncols):
                if sheet.cell_value(0, i) == col_name:
                    index = i
                    break
            for i in range(sheet.nrows):
                ctype = sheet.cell(i, index).ctype
                cell_value = sheet.cell_value(i, index)
                if ctype == 2 and cell_value % 1 == 0.0:  # ctype为2且为浮点
                    cell = int(cell_value)  # 浮点转成整型
                else:
                    cell = cell_value
                if value == str(cell):
                    res_list.append(i)
        if '不是' == operator:
            for i in range(sheet.ncols):
                if sheet.cell_value(0, i) == col_name:
                    index = i
                    break
            for i in range(sheet.nrows):
                ctype = sheet.cell(i, index).ctype
                cell_value = sheet.cell_value(i, index)
                if ctype == 2 and cell_value % 1 == 0.0:  # ctype为2且为浮点
                    cell = int(cell_value)  # 浮点转成整型
                else:
                    cell = cell_value
                if value != str(cell):
                    res_list.append(i)
        if '包含' == operator:
            for i in range(sheet.ncols):
                if sheet.cell_value(0, i) == col_name:
                    index = i
                    break
            for i in range(sheet.nrows):
                cell_value = sheet.cell_value(i, index)
                if i == 0 or str(value) in str(cell_value):
                    res_list.append(i)
        return res_list

    @staticmethod
    def get_values_by_col_name(sheet, col_name):
        index = 0
        res_list = []
        # 找到对应列名的列
        for i in range(sheet.ncols):
            if sheet.cell_value(0, i) == col_name:
                index = i
        for i in range(sheet.nrows):
            res_list.append(sheet.cell_value(i, index))
        return res_list

    def get_values_by_coordinate(self, rows_, cols_):
        two = []
        for row in rows_:
            one = []
            for col in cols_:
                one.append(self.mainsheet.cell_value(row, col))
            two.append(one)
        return two

    def get_cols_by_col_names(self, rows_name):
        cols_ = []
        for name in rows_name:
            for i in range(self.mainsheet.ncols):
                if self.mainsheet.cell_value(0, i) == name:
                    cols_.append(i)
        return cols_

    def get_row_by_index(self, index):
        res = []
        for i in range(self.mainsheet.ncols):
            res.append(self.mainsheet.cell_value(index, i))
        return res

    @staticmethod
    def write_excel(values_, excel_name_):
        wt_book = xlwt.Workbook()
        new_sheet = wt_book.add_sheet(excel_name_)
        # two = []
        for i in range(len(values_)):
            # one = []
            for j in range(len(values_[0])):
                # one.append((i, j))
                new_sheet.write(i, j, values_[i][j])
            # two.append(one)
        return wt_book

    @staticmethod
    def save_book(book_):

        desk = os.path.join(os.path.expanduser("~"), 'Desktop') + '\\'

        # key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
        #                      r'Software\Microsoft\Windows\CurrentVersion\Explorer\ShellFolders', )

        book_.save(desk + "新的工作簿.xls")
