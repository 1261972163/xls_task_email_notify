# -*- coding: utf-8 -*-
"""
Created on Sat Feb 17 12:14:44 2018

@author: yangleyuan
"""
#0.准备工作，定义二个日志输出：runlog.log与console
import logging
# 配置日志信息
logging.basicConfig(level=logging.DEBUG,
          format='%(asctime)s %(name)-12s %(levelname)-8s %(message)s',
          datefmt='%m-%d %H:%M',
          filename='runlog.log',
          filemode='w')
# 定义一个Handler打印INFO及以上级别的日志到sys.stderr
console = logging.StreamHandler()
console.setLevel(logging.INFO)
# 设置日志打印格式
formatter = logging.Formatter('%(name)-12s: %(levelname)-8s %(message)s')
console.setFormatter(formatter)
# 将定义好的console日志handler添加到root logger
logging.getLogger('').addHandler(console)

logging.debug('debug message')
logging.info('info message')
logging.warning('warning message')
logging.error('error message')
logging.critical('critical message')


#1.读excel文件,并存入数据结构中提供数据访问接口
#读excel文件
def read_excel(excelFile, mama_task_list):
    import xlrd  #安装读excel的文件工具包
    data = xlrd.open_workbook(excelFile)
    table = data.sheets()[0]
    nrows = table.nrows #行数
    ncols = table.ncols #列数
    start = False
    for i in range(0,nrows):
        Flag = False
        rowValues= table.row_values(i) #某一行数据
#        print(rowValues)
        if start ==False and 'Title' in rowValues:
            start = True  #若此行元素item有'Title',则开始记录有效行

        for item in rowValues:
            if '' != item:
                #若此行元素item有非空元素，则把这一行置为有效行
                Flag = True

        if start and Flag:
            #若此行元素item有非空元素，则把这一行添加到mama_task_list
            mama_task_list.append(rowValues)

#数据访问接口
#数据字段
field_list=['Title', 'Detail', 'Date', 'State', 'Receiver', 'Cc', 'Rule', 'priority']
#创建数据字段索引表
def create_fields_index(mama_task_list):
    field_index = {}
    for i in field_list:
        index = mama_task_list[0].index(i)
        field_index.setdefault(i,index)
#    print(field_index)
    logging.debug('field_index：={}'.format(field_index))
    return field_index
#创建第k个任务记录
def create_task_record(mama_task_list, k, field_index):
    mama_task_record = {}
    for i in field_list:
        field_value = mama_task_list[k][field_index[i]]
        mama_task_record.setdefault(i,field_value)
#    print(mama_task_record)
#    print('================')
    logging.debug('mama_task_record={}'.format(mama_task_record))
    return mama_task_record
#创建整个任务数据表
def create_task_table(mama_task_list, field_index):
    mama_task_table =  []
    # 每一行任务信息处理
    for i in range(1,len(mama_task_list)):
        mama_task_record=create_task_record(mama_task_list, i, field_index)
        mama_task_table.append(mama_task_record)
#    print(mama_task_table)
#    print('================')
    logging.debug('mama_task_table={}'.format(mama_task_table))
    return mama_task_table

#2.发邮件
def mail(subject, receiver, cc, detail):
    import datetime
    import smtplib  #加载smtplib模块
    from email.mime.text import MIMEText
    from email.utils import formataddr
    my_sender='XXXXXXXXXXXXX@XXXXXX.com' #发件人邮箱账号，为了后面易于维护，所以写成了变量
    my_smtp='smtp.163.com' #发件人邮箱中的SMTP服务器
    my_smtp_port=25  #发件人邮箱中的SMTP服务器端口号
    my_passwd="XXXXXXXXXXXXXXXX" #发件人邮箱密码
    ret=True
    try:
        msg=MIMEText(detail,'plain','utf-8') #'plain'/'html' utf-8支持中文
        msg['From']=formataddr(["邮件提醒小助手",my_sender])   #括号里的对应发件人邮箱昵称、发件人邮箱账号
        msg['To']=receiver #收件人邮箱账号
        msg['Cc']=cc #抄送人邮箱账号
        msg['Subject']=subject #邮件的主题，也可以说是标题

        server=smtplib.SMTP(my_smtp,my_smtp_port)  #发件人邮箱中的SMTP服务器，端口是25
        server.login(my_sender,my_passwd)    #括号中对应的是发件人邮箱账号、邮箱密码

        toaddr=receiver.split(',') + cc.split(',')
        server.sendmail(my_sender,toaddr,msg.as_string())   #括号中对应的是发件人邮箱账号、收件人邮箱账号、发送邮件

        server.quit()   #这句是关闭连接的意思

    except Exception:   #如果try中的语句没有执行，则会执行下面的ret=False
        ret=False

    if ret:
#        print("[{}] Send email ok:({},{},{})".format(datetime.datetime.now(), subject,toaddr,detail)) #如果发送成功则会返回ok，稍等20秒左右就可以收到邮件
        logging.info("[{}] Send email ok:({},{},{})".format(datetime.datetime.now(), subject,toaddr,detail))
    else:
        print("[{}] Send email failure:({},{},{})".format(datetime.datetime.now(), subject,toaddr,detail))  #如果发送失败则会返回failure
        logging.error("[{}] Send email failure:({},{},{})".format(datetime.datetime.now(), subject,toaddr,detail))
    return ret

#3.读取excel一次性批量发送提醒邮件任务,实现简单，资源消耗小
def batch_mail_notify(mama_task_table):

    import datetime

    #根据任务信息逐个发送提醒邮件
    for i in range(0,len(mama_task_table)):
        try:
            #取任务表中字段
            title=mama_task_table[i]['Title']
            detail=mama_task_table[i]['Detail']
            date=mama_task_table[i]['Date']
            receiver=mama_task_table[i]['Receiver']
            cc = mama_task_table[i]['Cc']
            rule = mama_task_table[i]['Rule']
            #计算年月日时分秒，excel时间定义格式要标准化，否则会出错
            date1 = datetime.datetime.strptime(date,"%Y-%m-%d %H:%M:%S")
            #对于已经过去的任务不再提醒，只提醒还未过时的&今天的任务
            now = datetime.datetime.now()
            if date1 >= now:
                if [date1.year,date1.month,date1.day] == [now.year,now.month,now.day]:
                    mail(subject=title,receiver=receiver,cc=cc, detail=detail)
            logging.info("task [{}] process success！！！".format(i))

        except Exception:   #如果try中的语句执行出现异常
            print("task [{}] process failure！！！".format(i))  #如果发送失败则会返回failure
            logging.error("task [{}] process failure！！！".format(i))


#3.读取excel，根据每个任务时间单独起一个定时器任务去执行邮件提醒
def alone_mail_notify(mama_task_table):

    #根据任务信息逐个发送提醒邮件
    import datetime
    from threading import Timer

    global timer_list
    timer_list=[]  #这个用于保存每个任务的定时器任务列表，用于以后特殊情况下取消定时器用

    # 再根据每一行任务信息发送提醒邮件
    for i in range(0,len(mama_task_table)):
        try:
            #取任务表中字段
            title=mama_task_table[i]['Title']
            detail=mama_task_table[i]['Detail']
            date=mama_task_table[i]['Date']
            receiver=mama_task_table[i]['Receiver']
            cc = mama_task_table[i]['Cc']
            rule = mama_task_table[i]['Rule']

            #计算年月日时分秒，excel时间定义格式要标准化，否则会出错
            date1 = datetime.datetime.strptime(date,"%Y-%m-%d %H:%M:%S")
            #对于已经过去的任务不再提醒，只提醒还未过时的任务。（这个方式无需限定是今天的任务）
            if date1 >= datetime.datetime.now():
                d = date1 - datetime.datetime.now()
                delay = d.days*24*60*60+d.seconds
                logging.debug("d={};delay={};rule={}".format(d, delay, rule*60))
                if delay >= rule*60:
                    delay = delay - rule*60
                else:
                    delay = 0
                timer = Timer(delay,  mail, (title,receiver, cc, detail,))
                timer_list.append(timer)
                timer.start() #非阻塞型
                logging.debug("Timer start:=[{},{},{},{}]".format(delay,title,receiver, cc))
    #            print(timer)
            logging.info("task [{}] process success！！！".format(i))

        except Exception:   #如果try中的语句执行出现异常
            print("task [{}] process failure！！！".format(i))  #如果发送失败则会返回failure
            logging.error("task [{}] process failure！！！".format(i))

#    print(timer_list)
    logging.debug("timer_list:={}".format(timer_list))
    #将任务定时器保存进文件

#4.周期任务：简化可以每天00：00：00批量执行一次发送提醒邮件任务，体验差一些，但资源消耗比较小
# 定时任务：更合理做法，是根据每个任务时间单独起一个定时器去执行，体验好但资源消耗大。
def mail_Task():
    print('===================')

if __name__ == "__main__":
    #读取excel文件，获取任务信息
    mama_task_list=[]  #任务信息列表，用于保存整个excel所有行的任务数据
    excelFile  = './to_do_list.xlsx'
    read_excel(excelFile,mama_task_list)
#    print(mama_task_list)
#    print('================')
    logging.debug('mama_task_list={}'.format(mama_task_list))

    # 创建任务字段位置的索引表
    field_index=create_fields_index(mama_task_list)

    # 创建任务数据表
    mama_task_table=create_task_table(mama_task_list, field_index)

    #判断有无任务定时器文件，若无则创建；否则需要读取文件并将所有任务取消掉，重置所有任务定时器
    #由于定时器无法序列化进行保存
    #猜测：定时器是否与此程序绑定，在此程序结束时会自动清楚吗？？？若此则无需保存去取消了。
#    alone_mail_notify(mama_task_table)
    batch_mail_notify(mama_task_table)