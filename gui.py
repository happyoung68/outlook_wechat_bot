import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLineEdit, QPushButton, QLabel, QMessageBox
import win32com.client
import os
import re
import pandas as pd
import requests
from datetime import datetime, timedelta
import textwrap




def read_outlook_emails(message_name, shaixuan_name):
    current_dir = os.getcwd()
    # 连接 Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # 获取收件箱
    inbox = outlook.GetDefaultFolder(6)  # 6 表示收件箱

    # 获取邮件列表（按接收时间倒序）
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    # 遍历邮件
    for i, msg in enumerate(messages):
        if i >= 10:
            break
        if msg.Subject == message_name:
            print(f"\n=== 邮件 {i + 1} ===")
            print(f"主题: {msg.Subject}")
            print(f"发件人: {msg.SenderName}")
            print(f"时间: {msg.ReceivedTime}")
            print(f"正文: {msg.Body}")
            for attachment in msg.Attachments:
                # 获取附件文件名
                file_name = attachment.FileName
                print(f"附件名: {file_name}")
                save_path = os.path.join(current_dir, file_name)
                # 检查附件是否为 .xlsx 格式
                if file_name.lower().endswith('.xlsx'):
                    # 保存附件到本地
                    attachment.SaveAsFile(save_path)
                    print(f"已保存附件: {save_path}")
                    return msg.Body, save_path
                else:
                    print(f"已忽略附件: {file_name}")

class MyWindow(QWidget):
    def __init__(self):
        super().__init__()

        # 设置窗口的基本属性
        self.setWindowTitle("SOA发送机器人")
        self.setGeometry(500, 300, 400, 200)

        # 创建控件
        self.layout = QVBoxLayout()

        self.input_var1 = QLineEdit(self)
        self.input_var1.setPlaceholderText("请输入邮件名,示例：xxx")
        self.layout.addWidget(self.input_var1)

        self.input_var2 = QLineEdit(self)
        self.input_var2.setPlaceholderText("请输入筛选名，示例：xxxx")
        self.layout.addWidget(self.input_var2)

        self.run_button = QPushButton("运行", self)
        self.run_button.clicked.connect(self.run_function)  # 点击时运行函数
        self.layout.addWidget(self.run_button)

        self.result_label = QLabel("结果会显示在这里", self)
        self.layout.addWidget(self.result_label)

        self.setLayout(self.layout)

    def run_function(self):

        # 获取输入框中的值
        message_name = str(self.input_var1.text()).strip()  # 去掉两端的空格
        shaixuan_name = str(self.input_var2.text()).strip()  # 去掉两端的空格

        # 检查输入框是否为空
        if not message_name or not shaixuan_name:
            QMessageBox.warning(self, "输入错误", "请输入邮件名和筛选关键词", QMessageBox.Ok)
            return

        zhengwen, save_path = read_outlook_emails(message_name, shaixuan_name)
        # print(zhengwen)
        gi_pattern = r"GI版本[^】]*【([^】]+)】"
        host_pattern = r"HOST[^】]*【([^】]+)】"
        mproxy_pattern = r"V[0-9]+\.[0-9]+\.[0-9]+_[0-9]{8}"

        # 提取版本号
        gi_version = re.search(gi_pattern, zhengwen)
        host_version = re.search(host_pattern, zhengwen)
        mproxy_version = re.search(mproxy_pattern, zhengwen)

        # 输出提取结果
        if gi_version:
            print(f"GI版本号: {gi_version.group(1)}")
        else:
            print("未找到GI版本号")

        if host_version:
            print(f"Host版本号: {host_version.group(1)}")
        else:
            print("未找到Host版本号")

        if mproxy_version:
            print(f"MProxy版本号: {mproxy_version.group(0)}")
        else:
            print("未找到MProxy版本号")

        # 读取 Excel 文件
        sheet_name = "ServiceInterfaces"  # 你要读取的工作表名称
        df = pd.read_excel(save_path, sheet_name=sheet_name)  # header=[0, 1] 让 pandas 读取合并的表头
        column_cnt = 0
        for column, value in df.iloc[0].items():
            if value == "临时打N":
                break
            column_cnt += 1

            # print(f"列名: {column}, 单元格内容: {value}")
        print("写临时打N的是第几列：", column_cnt)
        df = df[df.iloc[:, column_cnt] == shaixuan_name]
        # # 查看列名，确认列名
        # print(df.columns)

        # 假设'AI'列是第35列（索引从0开始），'C'列是第2列（假设你知道列的位置）
        responsible_persons = df.iloc[:, 34]  # 使用索引来获取负责人列
        functions = df.iloc[:, 0]  # 使用索引来获取功能列
        # print(functions)
        # print(responsible_persons)

        # 创建一个字典来存储每个负责人的功能
        responsibility_dict = {}

        # 遍历数据并整理
        for person, function in zip(responsible_persons, functions):
            if person not in responsibility_dict:
                responsibility_dict[person] = set()  # 使用 set 去重功能
            responsibility_dict[person].add(function)

        # 输出每个负责人的功能总结
        # for person, functions in responsibility_dict.items():
        #     print(f"1、{'、'.join(map(str, functions))}——@{person}")

        for index, (person, functions) in enumerate(responsibility_dict.items(), start=1):
            print(f"{index}、{'、'.join(map(str, functions))}——@{person}")

        # 生成人员和功能的 Markdown 列表
        responsibility_list = "\n".join([f"{index}. {'、'.join(map(str, functions))}——@{person}   "
                                         for index, (person, functions) in
                                         enumerate(responsibility_dict.items(), start=1)])

        current_time = datetime.now()
        time_plus_one_day = current_time + timedelta(days=1)
        # 配置区（需修改部分）
        WEBHOOK_URL = "xxxxxxxxxxx"
        markdown_content = f"""
🚨 **【MID_SOA编译告警】需1个工作日内确认**  

▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂  
■ **失效版本**  
GI版本：`{gi_version.group(1)}`
Host版本：`{host_version.group(1)}`  
Mproxy矩阵：`{mproxy_version.group(0)}`  

■ **失效服务**   
{responsibility_list} 
▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂  
**‼️ 处理方式（任选其一）**  
-选项1：矩阵错误，实际该车型并不需要该服务，会尽快发起矩阵变更申请删除相关服务；  
-选项2：应用接口错误或缺失，要求SOA配合在本阶段修复实现该服务；  
<font color="info">（服务owner行动项：①说明接口错误或缺失的原因；②说明在本阶段实现的必要性；③向对应项目DRE申请修复软件版本；④确保应用在对应版本修复相关接口。）</font>  
-选项3：应用接口错误或缺失，但在本轮装车阶段`{gi_version.group(1)}`不再实施，不再需要编译该服务，留待下一轮才应对；  
-选项4：无法确认接口有何问题，要求SOA协同确认。  
▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂  
**⏰ 截止时间**  
<font color="warning">{time_plus_one_day}</font>  
超时未反馈将关闭编译权限！  
            """

        data = {
            "msgtype": "markdown",
            "markdown": {"content": markdown_content}
        }

        requests.post(WEBHOOK_URL, json=data)
        self.result_label.setText(f"结果: \n{responsibility_list}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    sys.exit(app.exec_())
