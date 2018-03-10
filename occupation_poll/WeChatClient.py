import itchat


class WeChatClient:

    def __init__(self):
        self.itchat = itchat.auto_login(hotReload=True)
        self.friends = itchat.get_friends(update=True)[1:]

    def get_friends(self):
        friendList = itchat.get_friends(update=True)[1:]
        return friendList

    def get_classmates(filepath, num_range):
        pass

    def target(self):
        pass

    @itchat.msg_register(itchat.content.ATTACHMENT)
    def file_reply(self, msg):
        return_msg = "收到，谢谢！"
        return return_msg

    def get_friend(self, name):
        for friend in self.friends:
            if (friend["NickName"] == name) or (friend["RemarkName"] == name):
                return friend
        return None

    @itchat.msg_register(itchat.content.ATTACHMENT)
    def download_files(msg):
        msg.download('./data/' + msg.fileName)
        return_msg = "收到，谢谢！"
        return return_msg

    def send_msg(self, msg, user):
        itchat.send(msg, user)

    # @itchat.msg_register(itchat.content.ATTACHMENT)
    def send_file(self, file_dir, user):
        itchat.send_file(file_dir, toUserName=user)


    def start_service(self):
        itchat.run(True)
