import tkinter as tk
import poplib
import smtplib
from email.mime.text import MIMEText
from email.parser import Parser
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from tkinter import filedialog
import os

# POP3服务器信息
pop_server = 'pop.qq.com'
pop_port = 995
pop_username = '1961881107@qq.com'
pop_password = 'tjprmxbljctsecie'

# SMTP服务器信息
smtp_server = 'smtp.qq.com'
smtp_port = 587
smtp_username = '1961881107@qq.com'
smtp_password = 'tjprmxbljctsecie'

email_contents = {}  # 在这里定义一个字典来存储邮件内容
attachments = []  # 用于存储附件路径的列表

# 创建邮件接收函数
# 修改接收邮件函数
def receive_email():
    try:
        # 连接到POP3服务器
        pop_conn = poplib.POP3_SSL(pop_server, pop_port)
        pop_conn.user(pop_username)
        pop_conn.pass_(pop_password)

        # 获取邮箱中的邮件信息
        email_count, email_size = pop_conn.stat()
        messages = [pop_conn.retr(i) for i in range(1, email_count + 1)]

        # 显示邮件主题列表
        email_list = tk.Tk()
        email_list.title("邮件列表")
        email_list.geometry("600x500")  # 调整窗口大小

        scrollbar = tk.Scrollbar(email_list)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        listbox = tk.Listbox(email_list, yscrollcommand=scrollbar.set, width=85,height=25)  # 增大Listbox高度
        listbox.pack()

        
        for i, msg in enumerate(messages, 1):
            # 解码邮件内容
            content = b'\n'.join(msg[1]).decode('utf-8')
            parsed_msg = Parser().parsestr(content)
            sender = parsed_msg['from']
            send_time = parsed_msg['Date']
            
            # 只显示发送者和发送时间
            listbox.insert(tk.END, f"({i}) From: {sender} Sent: {send_time}")
            
            # 存储邮件内容以备查看
            email_contents[i] = content  # 将整数作为键来存储邮件内容


        scrollbar.config(command=listbox.yview)

        def view_email():
            selected_index = listbox.curselection()[0]  # 获取用户选中的邮件索引
            selected_content = email_contents[selected_index + 1]  # 获取选中邮件的内容（使用整数索引）
            show_email(selected_content)


        def show_email(content):
            email_window = tk.Toplevel()
            email_window.title("邮件内容")

            email_text = tk.Text(email_window)
            email_text.pack()

            parsed_msg = Parser().parsestr(content)
            email_body = ""

            # 寻找纯文本部分
            for part in parsed_msg.walk():
                if part.get_content_type() == 'text/plain':
                    email_body = part.get_payload(decode=True).decode(part.get_content_charset())
                    break

            email_text.insert(tk.END, email_body)


        view_button = tk.Button(email_list, text="查看邮件", command=view_email)
        view_button.pack()

        email_list.mainloop()

        # 关闭连接
        pop_conn.quit()
    except Exception as e:
        print("Error:", e)

# 创建邮件发送函数
# 创建邮件内容（包含附件）
def create_email():
    msg = MIMEMultipart()
    msg['Subject'] = subject_entry.get()
    msg['From'] = smtp_username
    msg['To'] = to_entry.get()

    body = message_entry.get("1.0", tk.END)
    msg.attach(MIMEText(body, 'plain'))

    # 添加附件
    for file_path in attachments:
        part = MIMEApplication(open(file_path, "rb").read())
        part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file_path))
        msg.attach(part)

    return msg


def browse_file():
    global attachments  # 确保引用全局变量attachments
    file_paths = filedialog.askopenfilenames()
    if file_paths:
        attachment_list.delete(0, tk.END)  # 清空附件列表
        attachments = list(file_paths)  # 更新attachments列表
        for file_path in file_paths:
            attachment_list.insert(tk.END, file_path)

def remove_attachment():
    selected_index = attachment_list.curselection()
    if selected_index:
        attachment_list.delete(selected_index)


# 创建邮件发送函数
def send_email():
    try:
        # 构建邮件内容
        msg = create_email()

        # 连接到SMTP服务器
        smtp_conn = smtplib.SMTP(smtp_server, smtp_port)
        smtp_conn.starttls()
        smtp_conn.login(smtp_username, smtp_password)

        # 发送邮件
        smtp_conn.sendmail(smtp_username, to_entry.get(), msg.as_string())

        # 关闭连接
        smtp_conn.quit()
        print("Email sent successfully!")
    except Exception as e:
        print("Error:", e)


# 创建界面
root = tk.Tk()
root.title("邮件客户端")

root.geometry("600x700")

# 发送邮件部分
to_label = tk.Label(root, text="收件人:")
to_label.pack()
to_entry = tk.Entry(root)
to_entry.pack()

subject_label = tk.Label(root, text="主题:")
subject_label.pack()
subject_entry = tk.Entry(root)
subject_entry.pack()

message_label = tk.Label(root, text="内容:")
message_label.pack()

# 设置内容文本框的宽度为80个字符，高度为20行
message_entry = tk.Text(root, width=80, height=20)
message_entry.pack()

# 接收邮件按钮
receive_button = tk.Button(root, text="接收邮件", command=receive_email)
receive_button.pack()

send_button = tk.Button(root, text="发送邮件", command=send_email)
send_button.pack()

attachment_button = tk.Button(root, text="选择附件", command=browse_file)
attachment_button.pack()

attachment_list = tk.Listbox(root, selectmode=tk.MULTIPLE)
attachment_list.pack()

remove_button = tk.Button(root, text="删除选定附件", command=remove_attachment)
remove_button.pack()

root.mainloop()
