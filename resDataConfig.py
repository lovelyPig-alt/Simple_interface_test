"""
测试结果Excel 字段配置
"""
from Utils.operationExcel import OperationExcel
class ResDataConfig:
    """
    测试结果Excel 字段配置   还需要为特殊的请求数据构造对应的请求内容
    """
    def __init__(self,file_name):
        self.opExcel = OperationExcel(file_name,0)
        self.test_id = 0  # 测试编号
        self.test_name = 1  # 测试名称
        self.request_url = 2  # 请求地址
        self.request_method = 3  # 请求方法 post get
        self.request_header = 4  # 请求头
        self.request_data = 5  # 请求数据
        self.note = 6  # 备注
        self.expected_res = 7  # 预期结果
        self.actual_res = 8  # 实际结果
        self.test_res = 9  # 测试结果

    def getTestIdcol(self):
        """
        获取测试编号的列数
        :return:
        """
        return self.test_id

    def getTestNamecol(self):
        """
        获取测试名称的列
        :return:
        """
        return self.test_name

    def getRequestUrlcol(self):
        """
        获取请求地址的列数
        :return:
        """
        return self.request_url

    def getRequestMethodcol(self):
        """
        获取请求方法的列数
        :return:
        """
        return self.request_method

    def getRequestHeadercol(self):
        """
        获取请求头的列数
        :return:
        """
        return self.request_header

    def getRequestDatacol(self):
        """
        获取请求数据的列数
        :return:
        """
        return self.request_data

    def getNotecol(self):
        """
        备注所在列
        :return:
        """
        return self.note

    def getExpectedResultcol(self):
        """
        获取预期结果的列数
        :return:
        """
        return self.expected_res

    def getActualResultcol(self):
        """
        获取实际结果的列数
        :return:
        """
        return self.actual_res

    def getTestResultcol(self):
        """
        测试结果的列
        :return:
        """
        return self.test_res


