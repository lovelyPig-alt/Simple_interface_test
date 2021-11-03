import xlrd
from xlutils.copy import copy  # 写入Excel

from Utils.operationExcel import OperationExcel
from Config.resDataConfig import ResDataConfig


class GetData:
    def __init__(self, file_name):
        # 初始化操作excel的对象
        self.opExcel = OperationExcel(file_name)
        self.res = ResDataConfig(file_name)

    def get_case_lines(self):
        """
        获取测试用例的个数 第一行为测试标题
        :return:
        """
        lines = self.opExcel.get_nrows() - 1
        return lines

    def get_test_name(self,row):
        """
        测试功能
        :return:
        """
        col = self.res.getTestNamecol()
        return self.opExcel.get_cel_value(row,col)

    def get_write_lines(self):
        """
        返回Excel中第一个空行
        :return:
        """
        return self.opExcel.get_nrows()

    def get_request_id(self, row):
        """
        测试用例id
        :param row:
        :return:
        """
        col = self.res.getTestIdcol()
        id = self.opExcel.get_cel_value(row, col)
        return str(id)

    def get_request_url(self, row):
        """
        获取请求地址
        :param row:
        :return:
        """
        col = self.res.getRequestUrlcol()
        request_url =self.opExcel.get_cel_value(row, col)
        return request_url

    def get_request_method(self, row):
        """
        请求方法
        :param row:
        :return:
        """
        col = self.res.getRequestMethodcol()
        request_method = self.opExcel.get_cel_value(row, col)
        return request_method

    def get_request_data(self, row):
        """
        请求数据 一般Post请求为json格式 get请求在url中携带
        :param row:
        :return:
        """
        col = self.res.getRequestDatacol()
        request_data = self.opExcel.get_cel_value(row, col)
        if request_data == "":
            return None
        return request_data


    def get_headers(self, row):
        """
        返回请求头
        :return:
        """
        col = self.res.getRequestHeadercol()
        headers = self.opExcel.get_cel_value(row, col)

        headers = eval(headers)
        # is_headers={
        #        'Authorization': 'Basic YWRtaW46MTIzNDU2', # admin:123456
        #     }
        return headers

    def get_note(self,row):
        """
        获取备注
        :param row:
        :return:
        """
        col = self.res.getNotecol()
        note = self.opExcel.get_cel_value(row,col)
        return note

    def get_expected_data(self, row):
        """
        获取预期结果
        :param row:
        :return:
        """
        col = self.res.getExpectedResultcol()
        expected_data = self.opExcel.get_cel_value(row, col)
        return expected_data

    def get_test_reuslt(self,row):
        """
        获取测试结果
        :param row:
        :return:
        """
        col = self.res.getTestResultcol()
        test_res  = self.opExcel.get_cel_value(row,col)
        return test_res

    def get_actual_data(self, row):
        """
        获取实际结果
        :param row:
        :return:
        """
        col = self.res.getActualResultcol()
        actual_data = self.opExcel.get_cel_value(row, col)
        return actual_data

    def write_actual_value(self, row, value):
        """
        写入实际结果
        :param row:
        :param value:
        :return:
        """
        col = self.res.getActualResultcol()
        work_book = xlrd.open_workbook(self.opExcel.file_name, formatting_info=True)
        # 先通过xlutils.copy下copy复制Excel
        write_to_work = copy(work_book)
        # 通过sheet_by_index没有write方法 而get_sheet有write方法
        sheet_data = write_to_work.get_sheet(self.opExcel.sheet_id)
        sheet_data.write(row, col, str(value))
        # 这里要注意保存 可是会将原来的Excel覆盖 样式消失
        write_to_work.save(self.opExcel.file_name)

    def write_test_res(self, row, value):
        """
        写入测试结果
        :param row:
        :param value:
        :return:
        """
        col = self.res.getTestResultcol()
        work_book = xlrd.open_workbook(self.opExcel.file_name, formatting_info=True)
        write_to_work = copy(work_book)
        sheet_data = write_to_work.get_sheet(self.opExcel.sheet_id)
        sheet_data.write(row, col, str(value))
        write_to_work.save(self.opExcel.file_name)

    def write_all_data(self, row, init_id, name, url, method, headers, data, note, predicted_code, response,
                       testresult):
        """
        写入全部数据
        :param row:   写入行号
        :param init_id:   初始iD
        :param name:   测试功能
        :param url:    接口地址
        :param method:  请求方法
        :param data:    请求数据
        :param headers: 请求头
        :param note:     备注
        :param predicted_code   预测状态码:
        :param response:         结果
        :param testresult:       测试结果
        :return:
        """
        tem = [init_id, name, url, method, headers, data, note, predicted_code, response, testresult]
        work_book = xlrd.open_workbook(self.opExcel.file_name, formatting_info=True)
        write_to_work = copy(work_book)
        sheet_data = write_to_work.get_sheet(self.opExcel.sheet_id)
        for col, val in enumerate(tem):
            if isinstance(val, float):
                val = str(int(val))
            if isinstance(val, list):
                val = " : ".join(list(map(str, val)))

            sheet_data.write(row, col, val)
        write_to_work.save(self.opExcel.file_name)

