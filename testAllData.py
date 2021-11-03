import os


from Utils.getResData import GetData

from Utils.runMain import RunMain
from Utils.utils import Utils


class FunctionTest:
    def __init__(self, res_excel):
        self.res_excel=res_excel
        self.runMethod = RunMain()
        self.res_data = GetData(self.res_excel)
        self.util = Utils()
        self.toekn  = " "
        self.memberNo = " "
        self.shuoshuoNo = " "
        self.praiseTypeNo = " "


    def run_test(self):
        """
        获取Excel中全部测试用例在对应接口进行测试
        :return:
        """
        # 获取测试用例个数
        lines = self.res_data.get_case_lines()

        #第0行为标题
        for row in range(1,lines+1):
            name = self.res_data.get_test_name(row)  # 测试功能
            url = self.res_data.get_request_url(row)  # 测试接口
            method = self.res_data.get_request_method(row)  # 测试方法
            data = self.res_data.get_request_data(row)  # 所有测试用例
            header = self.res_data.get_headers(row)
            note = self.res_data.get_note(row)  # 测试备注
            statecode = self.res_data.get_expected_data(row)  # 状态码

            data = eval(data)
            if "token" in data:  # 如果该接口需要token
                data['token'] = self.toekn
            '''
                # 如果接口需要会员编号
            if "memberNo" in data:
                data['memberNo'] = self.memberNo
            if "shuoshuoNo" in data:
                data["shuoshuoNo"] = self.shuoshuoNo
            '''
            # 真实响应
            response = self.runMethod.main(method, url, data, header)
            print("{}:正在测试：{},请求数据：{}，返回结果:{}".format(row,url,data,response))

            # 如果是登录接口更新token
            if url.split("/")[-1]=="appLogin":
                if response and response["errno"]==0:
                    self.toekn = response['data']['token']
            '''
            # 如果需要shuoshuoNo的接口，更新getShuoshuoList
            if url.split('/')[-1] == 'getShuoshuoList':
                if response and response['errno'] == 0:
                    self.shuoshuoNo = response['data']['rows'][0]["shuoshuoNo"]
            '''
            state = int(statecode)
            test_res = "Fail"
            if self.util.is_equal_code(state, response):
                test_res = "PASS"
            '''
            # 如果是需要memberNo的接口，更新memberNo
            if url.split('/')[-1] == 'appLogin':
                if response and response['status'] == 1:
                    self.memberNo = response['data']['memberNo']
            # 根据type传说说编号或者相册编号，更新getShuoShuoList或相册接口
            '''

            # 所有数据覆盖写入
            self.res_data.write_all_data(row, row, name, url, method, str(header), str(data), note, state, str(response), test_res)



if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="功能测试")
    parser.add_argument("-r", "--respath", help="测试结果",default="./Excel/res_excel/wkx_test.xls")
    args = vars(parser.parse_args())  # vars() 函数返回对象object的属性和属性值的字典对象。

    res_excels = args['respath']

    test = FunctionTest(res_excels)
    test.run_test()
