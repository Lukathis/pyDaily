from ExcelReader import ExcelReader
from WeChatClient import WeChatClient

import itchat
import os

if __name__=="__main__":

    # get classmates from document
    filepath = "."
    filename = "2015master.xls"
    er = ExcelReader(filepath, filename)

    # get all classmates' names
    classmates_range = range(22, 90)
    classmates_range = [90]

    classmates_name = er.read_column(classmates_range, 2)
    name_num = er.read_two_column_as_pair(classmates_range, 2, 1)

    # get classmates' info from wechat
    client = WeChatClient()
    friends = client.get_friends()
    classmates = {}

    for classmate_name in classmates_name:
        classmate = client.get_friend(classmate_name)
        if classmate != None:
            classmates[classmate_name] = classmate
    # print(classmates)
    # print(len(classmates))

    # send files to classmates
    msg = "打扰了，班主任需要统计一下毕业去向，麻烦填一下表里最后两个格子再发还给我就可以了，谢谢！"
    for classmate_name in classmates_name:
        # send message
        client.send_msg(msg, classmates[classmate_name]['UserName'])

        # send files
        blank_excel = name_num[classmate_name] + ".xls"
        blanck_excel_path = os.path.join("./blank", blank_excel)
        client.send_file(blanck_excel_path, classmates[classmate_name]['UserName'])
        print(blanck_excel_path)

        print(classmate)
        sleep(5)

    client.start_service()


    


