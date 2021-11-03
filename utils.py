import base64



class Utils:
    """
    工具类
    """

    def __init__(self):
        pass

    def str_is_equal(self, str1, str2):
        """
        判断两个字符串变为字典是否相等
        :param str1:
        :param str2:
        :return: bool
        """
        str1 = str1.replace("true", "True")
        if str1 == "":
            return False
        if not isinstance(str1, dict):
            str1 = eval(str1)
        if not isinstance(str2, dict):
            str2 = eval(str2)

        try:
            return str1['returnState']['stateCode'] == str2['returnState']['stateCode']
        except Exception as e:
            print("Error")
            return False

    def is_equal_code(self, code, res):
        """
        检测指定状态码和返回结果状态码是否一致
        :param code:
        :param res:
        :return:
        """
        if not isinstance(res, dict):
            res = eval(res)
        return res['status'] == code








