# exchange

#亲测可用，但是会发送，不会只保存在草稿箱
from exchangelib import DELEGATE, Account, Credentials, Configuration, Build, Version, Message, Mailbox, HTMLBody,NTLM,IMPERSONATION,FileAttachment
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
import urllib3
import os
import pandas as pd
#忽略警告
urllib3.disable_warnings()

#要是一直有max retries exceeded with url,一定要加这句
version = Version(build=Build(15,0,12,34))

#此句用来消除ssl证书错误，exchange使用自签证书需加上
BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter


#recipient_email=  "madanrong@szunicom.com"
#subject=u'草稿箱测试邮件'
#content = '测试存在草稿箱'
def save_email_draft(sender,recipient,subject,content,attachmentpath,cc_email):
    a = Account(
        primary_smtp_address=sender,
        config=config,
        autodiscover=False,
        access_type=DELEGATE,
    )
    recipient_email = recipient
    m = Message(
        account=a,
        folder=a.sent,
        subject=subject,
        body=HTMLBody(content),
        to_recipients=[Mailbox(email_address=i) for i  in recipient_email],
        cc_recipients = cc_email
    )
    for attachmentfile in attachmentpath:
        with open(attachmentfile,'rb') as file:
            attachment = FileAttachment(name = os.path.basename(attachmentfile),content = file.read())
            m.attach(attachment)
        #发送多人用[Mailbox(email_address='***@*****.com'),Mailbox(email_address='***@*****.com')...]
    m.save()
   # m.send_and_save() #发送
    m.move(a.drafts)
def basicinfo(path,info):
    data = pd.read_excel(path + info, sheet_name=0)
    num  = data.shape[0]
    recipient_email = data['主送'].tolist()
    subjects = data['主题'].tolist()
    leader  = data['抄送'].tolist()
    leaders = []
    recipient_emails =[]
    for i in range(num):
        recipient_emails.append (recipient_email[i].split(";") )
    for i in range(num):
        leaders.append (leader[i].split(";") )
    contents = data['邮件内容'].tolist()
    attachmentpath =  data['附件'].tolist()
    attachmentpaths =[]
    for i in range(num):
        attachmentpaths.append (attachmentpath[i].split(";") )
    return num,recipient_emails,leaders,contents,attachmentpaths,subjects
def senderinfo(path,info):
    data = pd.read_excel(path + info, sheet_name=1)
    NAME = data['账号'].tolist()
    sender = data['邮箱'].tolist()
    pwd = data['密码'].tolist()
    szunicom  = data['公司邮箱服务器'].tolist()
    return sender[0],pwd[0],NAME[0],szunicom[0]

path = os.getcwd()
info = r'\info.xlsx'
# 输入你的域账号如example\leo
sender,pwd,NAME,szunicom =senderinfo(path,info)
cred = Credentials(username= sender, password=pwd)
#attachmentpath = r'C:\Users\unicom\Downloads\深圳月报模板 (8).xlsx'

config = Configuration(service_endpoint=szunicom, credentials=cred, version=version, auth_type=NTLM)

sender= sender


num,recipient_emails,leaders,contents,attachmentpaths,subjects =basicinfo(path,info)
#print(leaders)
#i  = 1
#print(sender, recipient_emails[i], subjects[i], contents[i], attachmentpaths[i])
#save_email_draft(sender, sender, subjects[i], contents[i], attachmentpaths[i],cc)


for i in range(num):
    save_email_draft(sender, recipient_emails[i], subjects[i], contents[i], attachmentpaths[i],leaders[i])
    print("写给{}的邮件存储好了".format(recipient_emails[i]))
