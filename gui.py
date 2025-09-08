import sys
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QLineEdit, QPushButton, QLabel,
                             QMessageBox, QHBoxLayout, QGroupBox, QTextEdit, QFrame, QProgressBar)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QPalette, QColor, QIcon
import win32com.client
import os
import re
import requests
from datetime import datetime, timedelta
from PyQt5.QtGui import QFont, QPalette, QColor, QPixmap, QIcon, QTextCursor
from openpyxl import load_workbook
import pandas as pd
from bs4 import BeautifulSoup
import zipfile
import win32com.client as win32


class MyWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SOA 服务告警机器人")
        self.setGeometry(500, 300, 800, 600)

        # 设置应用样式
        self.setStyleSheet("""
            QWidget {
                background-color: #2c3e50;
                color: #ecf0f1;
                font-family: 'Segoe UI', '微软雅黑', sans-serif;
            }
            QGroupBox {
                background-color: #34495e;
                border: 2px solid #3498db;
                border-radius: 10px;
                margin-top: 15px;
                padding: 15px;
                font-weight: bold;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top center;
                padding: 0 10px;
                background-color: #3498db;
                color: white;
                border-radius: 5px;
            }
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px 20px;
                font-weight: bold;
                font-size: 14px;
                min-height: 40px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:pressed {
                background-color: #1c6ea4;
            }
            QPushButton:disabled {
                background-color: #7f8c8d;
            }
            QLabel {
                color: #ecf0f1;
                font-size: 14px;
            }
            QLineEdit {
                background-color: #2c3e50;
                color: #ecf0f1;
                border: 2px solid #3498db;
                border-radius: 6px;
                padding: 8px;
                font-size: 14px;
                selection-background-color: #3498db;
            }
            QTextEdit {
                background-color: #2c3e50;
                color: #ecf0f1;
                border: 2px solid #3498db;
                border-radius: 6px;
                padding: 10px;
                font-size: 13px;
            }
            QFrame#divider {
                background-color: #3498db;
                max-height: 2px;
                min-height: 2px;
            }
            QProgressBar {
                border: 2px solid #3498db;
                border-radius: 5px;
                text-align: center;
                background-color: #2c3e50;
            }
            QProgressBar::chunk {
                background-color: #3498db;
                width: 10px;
            }
        """)

        # 主布局
        main_layout = QVBoxLayout()
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # 标题
        title_label = QLabel("SOA 服务告警机器人")
        title_label.setStyleSheet("""
            font-size: 24px;
            font-weight: bold;
            color: #3498db;
            padding-bottom: 15px;
        """)
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        # 输入参数区域
        input_group = QGroupBox("邮件参数设置")
        input_layout = QVBoxLayout(input_group)

        # 邮件名输入
        email_layout = QHBoxLayout()
        email_label = QLabel("邮件名称:")
        email_label.setMinimumWidth(100)
        self.input_var1 = QLineEdit()
        self.input_var1.setPlaceholderText("示例：[外部]T68-G G3 E.0 α1 N掉报错服务矩阵表")
        email_layout.addWidget(email_label)
        email_layout.addWidget(self.input_var1)
        input_layout.addLayout(email_layout)

        # 筛选名输入
        filter_layout = QHBoxLayout()
        filter_label = QLabel("筛选名称:")
        filter_label.setMinimumWidth(100)
        self.input_var2 = QLineEdit()
        self.input_var2.setPlaceholderText("示例：T68G_N_0502")
        filter_layout.addWidget(filter_label)
        filter_layout.addWidget(self.input_var2)
        input_layout.addLayout(filter_layout)

        # 装车阶段
        pt_layout = QHBoxLayout()
        pt_label = QLabel("装车阶段:")
        pt_label.setMinimumWidth(100)
        self.input_var3 = QLineEdit()
        self.input_var3.setPlaceholderText("示例：PT1")
        pt_layout.addWidget(pt_label)
        pt_layout.addWidget(self.input_var3)
        input_layout.addLayout(pt_layout)



        main_layout.addWidget(input_group)

        # 操作按钮区域
        button_layout = QHBoxLayout()

        self.preview_button = QPushButton("预览消息")
        self.preview_button.setIcon(QIcon.fromTheme("document-preview"))
        self.preview_button.clicked.connect(self.preview_function)
        button_layout.addWidget(self.preview_button)

        self.run_button = QPushButton("发送告警")
        self.run_button.setIcon(QIcon.fromTheme("mail-send"))
        self.run_button.clicked.connect(self.run_function)
        button_layout.addWidget(self.run_button)

        main_layout.addLayout(button_layout)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("就绪")
        main_layout.addWidget(self.progress_bar)

        # 结果预览区域
        result_group = QGroupBox("消息预览")
        result_layout = QVBoxLayout(result_group)

        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        result_layout.addWidget(self.result_text)

        main_layout.addWidget(result_group)

        # 状态栏
        self.status_label = QLabel("就绪 | 请设置邮件参数")
        self.status_label.setStyleSheet("""
            background-color: #34495e;
            padding: 8px;
            border-radius: 5px;
            font-size: 12px;
        """)
        main_layout.addWidget(self.status_label)

        self.setLayout(main_layout)

        # 初始化变量
        self.zhengwen = None
        self.save_path = None
        self.responsibility_dict = {}
        self.gi_version = ""
        self.host_version = ""
        self.mproxy_version = ""

        # 设置窗口图标
        self.setWindowIcon(self.create_icon())

    def create_icon(self):
        # 创建一个简单的程序图标（机器人）
        icon_pixmap = QPixmap(64, 64)
        icon_pixmap.fill(Qt.transparent)
        return QIcon(icon_pixmap)

    def update_progress(self, value, message):
        self.progress_bar.setValue(value)
        self.progress_bar.setFormat(message)
        QApplication.processEvents()

    # def read_outlook_emails(self, message_name, shaixuan_name):
    #     self.update_progress(10, "正在连接Outlook...")
    #     zhengwen = None
    #     xlsx_path = None
    #     current_dir = os.getcwd()
    #
    #     try:
    #         # 连接 Outlook
    #         outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    #
    #         # 获取收件箱
    #         inbox = outlook.GetDefaultFolder(6)  # 6 表示收件箱
    #
    #         # 获取邮件列表（按接收时间倒序）
    #         messages = inbox.Items
    #         messages.Sort("[ReceivedTime]", True)
    #
    #         self.update_progress(20, "正在搜索邮件...")
    #
    #         # 遍历邮件
    #         for i, msg in enumerate(messages):
    #             if i >= 50:  # 最多检查50封邮件
    #                 break
    #             if msg.Subject == message_name:
    #                 self.status_label.setText(f"找到匹配邮件: {msg.Subject}")
    #                 self.update_progress(30, "处理邮件内容...")
    #
    #                 # 处理正文
    #                 zhengwen = msg.Body
    #
    #                 # 处理附件
    #                 for attachment in msg.Attachments:
    #                     file_name = attachment.FileName
    #                     save_path = os.path.join(current_dir, file_name)
    #                     if file_name.lower().endswith('.xlsx'):
    #                         attachment.SaveAsFile(save_path)
    #                         xlsx_path = save_path
    #                         self.status_label.setText(f"已保存附件: {file_name}")
    #
    #                 if xlsx_path:
    #                     return zhengwen, xlsx_path
    #
    #         if not zhengwen or not xlsx_path:
    #             self.status_label.setText("未找到匹配的邮件或附件")
    #             return None, None
    #
    #     except Exception as e:
    #         self.status_label.setText(f"Outlook错误: {str(e)}")
    #         return None, None

    # def read_outlook_emails(self, message_name, shaixuan_name):
    #     self.update_progress(10, "正在连接Outlook...")
    #     zhengwen = None
    #     xlsx_path = None
    #     current_dir = os.getcwd()
    #
    #     try:
    #         # 连接 Outlook
    #         outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    #
    #         # 获取所有邮箱账户
    #         accounts = outlook.Folders
    #
    #         self.update_progress(20, "正在搜索邮件...")
    #
    #         # 遍历所有邮箱账户
    #         for account in accounts:
    #             print(account.Name)
    #             # 获取该账户下的所有文件夹
    #             for folder in account.Folders:
    #                 print(folder.Name)
    #                 # 找到收件箱
    #                 if folder.Name == "收件箱":
    #                     print(1)
    #                     inbox = folder
    #
    #                     # 获取邮件列表（按接收时间倒序）
    #                     messages = inbox.Items
    #                     messages.Sort("[ReceivedTime]", True)
    #
    #                     # 遍历邮件
    #                     for i, msg in enumerate(messages):
    #                         if i >= 50:  # 最多检查50封邮件
    #                             break
    #                         if msg.Subject == message_name:
    #                             self.status_label.setText(f"找到匹配邮件: {msg.Subject}")
    #                             self.update_progress(30, "处理邮件内容...")
    #
    #                             # 处理正文
    #                             zhengwen = msg.Body
    #
    #                             # 处理附件
    #                             for attachment in msg.Attachments:
    #                                 file_name = attachment.FileName
    #                                 save_path = os.path.join(current_dir, file_name)
    #                                 if file_name.lower().endswith('.xlsx'):
    #                                     attachment.SaveAsFile(save_path)
    #                                     xlsx_path = save_path
    #                                     self.status_label.setText(f"已保存附件: {file_name}")
    #
    #                             if xlsx_path:
    #                                 return zhengwen, xlsx_path
    #
    #         # 如果没有找到符合条件的邮件或附件
    #         if not zhengwen or not xlsx_path:
    #             self.status_label.setText("未找到匹配的邮件或附件")
    #             return None, None
    #
    #     except Exception as e:
    #         self.status_label.setText(f"Outlook错误: {str(e)}")
    #         return None, None

    def read_outlook_emails(self, message_name, shaixuan_name):
        self.update_progress(10, "正在连接Outlook...")
        zhengwen = None
        xlsx_path = None
        current_dir = os.getcwd()

        try:
            # 连接 Outlook
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

            # 获取所有邮箱账户
            accounts = outlook.Folders

            self.update_progress(20, "正在搜索邮件...")

            # 遍历所有邮箱账户
            for account in accounts:
                print(account.Name)
                # 获取每个账户的收件箱
                inbox = account.Folders.Item("收件箱")

                # 获取邮件列表（按接收时间倒序）
                messages = inbox.Items
                messages.Sort("[ReceivedTime]", True)

                # 遍历邮件
                for i, msg in enumerate(messages):
                    if i >= 50:  # 最多检查50封邮件
                        break
                    if msg.Subject == message_name:
                        self.status_label.setText(f"找到匹配邮件: {msg.Subject}")
                        self.update_progress(30, "处理邮件内容...")

                        # 处理正文
                        zhengwen = msg.Body

                        # 处理附件
                        for attachment in msg.Attachments:
                            file_name = attachment.FileName
                            save_path = os.path.join(current_dir, file_name)
                            if file_name.lower().endswith('.xlsx'):
                                attachment.SaveAsFile(save_path)
                                xlsx_path = save_path
                                self.status_label.setText(f"已保存附件: {file_name}")

                        if xlsx_path:
                            return zhengwen, xlsx_path

            # 如果没有找到符合条件的邮件或附件
            if not zhengwen or not xlsx_path:
                self.status_label.setText("未找到匹配的邮件或附件")
                return None, None

        except Exception as e:
            self.status_label.setText(f"Outlook错误: {str(e)}")
            return None, None

    def parse_email_content(self, zhengwen):
        self.update_progress(40, "解析邮件内容...")

        gi_pattern = r"GI版本[^】]*【([^】]+)】"
        host_pattern = r"HOST[^】]*【([^】]+)】"
        mproxy_pattern = r"V[0-9]+\.[0-9]+\.[0-9]+_[0-9]{8}"

        # 提取版本号
        self.gi_version = re.search(gi_pattern, zhengwen).group(1) if re.search(gi_pattern, zhengwen) else "未找到"
        self.host_version = re.search(host_pattern, zhengwen).group(1) if re.search(host_pattern, zhengwen) else "未找到"
        self.mproxy_version = re.search(mproxy_pattern, zhengwen).group(0) if re.search(mproxy_pattern,
                                                                                        zhengwen) else "未找到"

        return True

    def excel_remove_filter(path, sheet_name):
        # 启动 Excel
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False  # 后台运行

        try:
            # 打开工作簿
            wb = excel.Workbooks.Open(os.path.abspath(path))

            # 定位工作表
            for sheet in wb.Sheets:
                if sheet.Name == sheet_name:
                    # 清除筛选
                    if sheet.AutoFilterMode:
                        sheet.AutoFilterMode = False
                    # 清除特殊筛选
                    if sheet.FilterMode:
                        sheet.ShowAllData()
                    break

            wb.save(path)
            wb.Close()
            return 0
        finally:
            excel.Quit()

    def parse_excel_data(self, save_path, shaixuan_name):
        # save_path = "(07-28_编译反馈_AH8 G3 F.0 α1)GAC_VehicleA_SOMEIP_CMX_CCU_Phase1_(CCU_S32G_MProxy)_V2.2.26_202506031.xlsx"
        self.update_progress(50, "解析Excel数据...")
        # print(1)
        try:
            # print(2)
            # 读取 Excel 文件
            df = pd.read_excel(save_path, sheet_name="ServiceInterfaces")
            # print(3)
            # 查找列索引
            column_cnt = 0
            for column, value in df.iloc[0].items():
                if value == "临时打N":
                    print("临时打N列",column_cnt)
                    break
                column_cnt += 1

            column_owner_cnt = 0
            for column, value in df.iloc[0].items():
                if column == "Owner":
                    print("Owner列",column_owner_cnt)
                    break
                column_owner_cnt += 1

            # 筛选数据
            df = df[df.iloc[:, column_cnt] == shaixuan_name]
            responsible_persons = df.iloc[:, column_owner_cnt]
            functions = df.iloc[:, 0].astype(str) + '-->' + df.iloc[:, 5].astype(str)

            # 创建负责人-功能字典
            self.responsibility_dict = {}
            for person, function in zip(responsible_persons, functions):
                if person not in self.responsibility_dict:
                    self.responsibility_dict[person] = set()
                self.responsibility_dict[person].add(function)

            return True

        except Exception as e:
            self.status_label.setText(f"Excel解析错误: {str(e)}")
            return False

    def generate_markdown(self):
        self.update_progress(70, "生成告警消息...")

        # 生成人员和功能的 Markdown 列表
        responsibility_list = "\n".join(
            [f"{index}. {'、'.join(map(str, functions))}——<@{person}>   "
             for index, (person, functions) in enumerate(self.responsibility_dict.items(), start=1)]
        )

        current_time = datetime.now()
        time_plus_one_day = current_time + timedelta(days=1)

        markdown_content = f"""
🚨 **【MID_SOA编译告警】需1个工作日内确认**  

▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂  
■ **失效版本**  
- GI版本：**`{self.gi_version}`**  
- 装车版本：**`{str(self.input_var3.text()).strip()}`**  
- Host版本：`{self.host_version}`  
- Mproxy矩阵：`{self.mproxy_version}`  
■ **失效服务Service InterFace Name-->Element Name**   
{responsibility_list} 
▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂  
**‼️ 处理方式（任选其一）**  
1. 矩阵错误，实际该车型并不需要该服务，会尽快发起矩阵变更申请删除相关服务；  
2. 应用接口错误或缺失，要求SOA配合在本阶段修复实现该服务；  
<font color="info">（服务owner行动项：①说明接口错误或缺失的原因；②说明在本阶段实现的必要性；③向对应项目DRE申请修复软件版本；④确保应用在对应版本修复相关接口。）</font>  
3. 应用接口错误或缺失，但在本装车版本不再实施，不再需要编译该服务，留待下一轮才应对；  
4. 无法确认接口有何问题，要求SOA协同确认。  
▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂  
**⏰ 截止时间**  
<font color="warning">{time_plus_one_day.strftime('%Y-%m-%d %H:%M')}</font>  
超时未反馈将关闭编译权限！  
"""
        return markdown_content

    def preview_function(self):
        """预览功能，生成消息但不发送"""
        self.run_function(preview_only=True)

    def run_function(self, preview_only=False):
        """主运行函数，preview_only=True时只预览不发送"""
        # 重置状态
        self.progress_bar.setValue(0)
        self.result_text.clear()
        self.status_label.setText("处理中...")

        # 获取输入值
        message_name = str(self.input_var1.text()).strip()
        shaixuan_name = str(self.input_var2.text()).strip()
        pt_name = str(self.input_var3.text()).strip()

        # 验证输入
        if not message_name or not shaixuan_name:
            QMessageBox.warning(self, "输入错误", "请输入邮件名和筛选关键词", QMessageBox.Ok)
            self.status_label.setText("错误: 请输入邮件名和筛选关键词")
            return

        # 读取Outlook邮件
        self.zhengwen, self.save_path = self.read_outlook_emails(message_name, shaixuan_name)
        if not self.zhengwen or not self.save_path:
            self.status_label.setText("错误: 找不到有效的邮件或附件")
            return

        # 解析邮件内容
        if not self.parse_email_content(self.zhengwen):
            return

        # 解析Excel数据
        if not self.parse_excel_data(self.save_path, shaixuan_name):
            return

        # 生成Markdown内容
        markdown_content = self.generate_markdown()

        # 显示预览
        # self.result_text.setHtml(self.format_markdown_for_display(markdown_content))
        self.result_text.setText(markdown_content)
        self.update_progress(90, "生成预览完成")

        # 如果是预览模式，不发送消息
        if preview_only:
            self.status_label.setText("预览模式: 消息已生成但未发送")
            self.update_progress(100, "预览完成")
            return

        # 发送消息到企业微信
        self.update_progress(95, "发送告警消息...")
        self.send_to_wechat(markdown_content)

        self.status_label.setText("告警消息已发送")
        self.update_progress(100, "发送完成")

    def format_markdown_for_display(self, markdown_content):
        """将Markdown内容格式化为HTML用于显示"""
        html_content = markdown_content.replace("\n", "<br>")
        html_content = html_content.replace("▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂", "<hr>")
        html_content = html_content.replace("🚨", "<span style='color:#e74c3c; font-size:18px;'>🚨</span>")
        html_content = html_content.replace("■", "<span style='color:#3498db;'>■</span>")
        html_content = html_content.replace("‼️", "<span style='color:#e74c3c;'>‼️</span>")
        html_content = html_content.replace("⏰", "<span style='color:#f39c12;'>⏰</span>")

        # 添加基本样式
        styled_html = f"""
        <html>
            <head>
                <style>
                    body {{
                        font-family: 'Segoe UI', '微软雅黑', sans-serif;
                        color: #ecf0f1;
                        background-color: #2c3e50;
                        font-size: 14px;
                        line-height: 1.6;
                    }}
                    hr {{
                        border: 0;
                        height: 1px;
                        background: linear-gradient(to right, rgba(52, 152, 219, 0), rgba(52, 152, 219, 0.75), rgba(52, 152, 219, 0));
                        margin: 15px 0;
                    }}
                    strong {{
                        color: #3498db;
                    }}
                </style>
            </head>
            <body>
                {html_content}
            </body>
        </html>
        """
        return styled_html

    def send_to_wechat(self, markdown_content):
        """发送消息到企业微信"""
        try:
            WEBHOOK_URL = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=74287684-05e7-4536-bd49-d96504b8b835"  # MID_SOA播报
            # WEBHOOK_URL = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=775c679c-eb9d-4b44-b971-9f9369c2c32f"  # 测试机器人

            data = {
                "msgtype": "markdown",
                "markdown": {
                    "content": markdown_content
                }
            }

            response = requests.post(WEBHOOK_URL, json=data)
            if response.status_code == 200:
                self.status_label.setText("告警消息发送成功")
            else:
                self.status_label.setText(f"发送失败: {response.status_code}")

        except Exception as e:
            self.status_label.setText(f"发送错误: {str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    sys.exit(app.exec_())
