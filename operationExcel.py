import xlrd

class OperationExcel():
    """
    #以面向对象的方式操作Excel
    """

    def __init__(self, file_name, sheet_id=0):
        """
        初始化OperationExcel对象
        :param file_name:
        :param sheet_id: vv
        """
        self.file_name = file_name
        self.sheet_id = sheet_id
        self.tables = self.get_tables()

    def get_tables(self):
        """
        返回tables对象
        :return:
        """
        ecel = xlrd.open_workbook(self.file_name)
        tables = ecel.sheet_by_index(self.sheet_id)
        return tables

    def get_nrows(self):
        """
        获取表格行数
        :return:
        """
        return self.tables.nrows

    def get_ncols(self):
        """
        获取表格列数
        :return:
        """
        return self.tables.ncols

    def get_data_by_row(self, row):
        """
        根据行号获取某一行的内容
        :param row:
        :return:
        """
        if row < 0:
            row = 0
        if row > self.get_nrows():
            row = self.get_nrows()
        data = self.tables.row_values(row)
        return data

    def get_data_by_col(self, col):
        """
        根据列号返回某一列的内容
        :param col:
        :return:
        """
        if col < 0:
            col = 0
        if col > self.get_ncols():
            col = self.get_ncols()
        data = self.tables.col_values(col)
        return data

    def get_cel_value(self, row, col):
        """
        获取某个指定单元格的内容
        :param row:
        :param col:
        :return:
        """
        data = self.tables.cell_value(row, col)

        # ecxel中读取数据时默认将数字类型读取为浮点型
        if isinstance(data, float):
            data = int(data)
        return data
