import os
import shutil

from Config.testDataConfig import TestDataConfig

from Utils.getResData import GetData
from Utils.utils import Utils


class GenerateTestData:
    def __init__(self, test_excel, res_excel,template_excel,prefix_url):
        # 新建excel
        if not os.path.exists(res_excel):
            shutil.copy(template_excel,res_excel)

        self.test_excel=test_excel
        self.res_excel=res_excel
        self.prefix_url = prefix_url

        self.test_data = TestDataConfig(self.test_excel,self.prefix_url)
        self.res_data = GetData(self.res_excel)
        self.util = Utils()



    def generate_data(self):
        """
        获取Excel中全部测试用例  生成一个测试文件
        :return:
        """
        name = self.test_data.get_test_name()  # 测试功能
        url = self.test_data.get_test_url()  # 测试接口
        method = self.test_data.get_test_method()  # 测试方法

        # all_test = self.test_data.get_all_test()  # 所有测试用例


        headers  = []

        # 否则按照正常字典构建
        all_test = self.test_data.refactor_all_test()


        for i in range(len(all_test)):
            header={}
            headers.append(header)

        row_counts = len(all_test)  # 测试用例个数

        note = self.test_data.get_note()  # 测试备注
        statecode = self.test_data.get_statecode()  # 状态码

        # 在这里要解决的问题是 指定多个测试文件时该如何写入
        lines = self.res_data.get_write_lines()  # 获取Excel中第一个空行的行号

        for i in range(row_counts):

            state = int(statecode[i])  #状态码
            h = str(headers[i])              #请求头
            t  = str(all_test[i])       #请求数据
            test_res = ""               #测试结果
            n = note[i]                     #备注
            self.res_data.write_all_data(lines, lines, name, url, method,h, t,n,state, "", test_res)

            lines += 1


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="生成测试数据脚本")
    parser.add_argument("-f", "--root", help="测试文件夹根目录",default="./Excel/wkx_test/")
    parser.add_argument("-t", "--template_excel", help="测试结果模板",default="./Excel/res_excel/template_excel.xls")
    parser.add_argument("-r", "--respath", help="测试结果存放路径",default="./Excel/res_excel/wkx_test.xls")
    parser.add_argument("-p", "--prefix_url", help="服务器ip",default="http://dev.heiyiren.io/")
    args = vars(parser.parse_args())  # vars() 函数返回对象object的属性和属性值的字典对象。

    root_dir = args["root"]
    template_excel = args["template_excel"]
    res_excels = args['respath']
    prefix_url = args['prefix_url']
    # 定义测试的顺序



    all_task = ["hyr_login.xls","hyr_homeData.xls","hyr_depositTestnet.xls","hyr_config.countries.xls","hyr_config.appVersion.xls"
                ,'hyr_config.app.xls']

    for index,test_excel in enumerate(all_task):
            path = os.path.join(root_dir,test_excel)
            data = GenerateTestData(path, res_excels,template_excel,prefix_url)
            data.generate_data()
            print(" %s测试用例已生成 " % test_excel.split("/")[-1])