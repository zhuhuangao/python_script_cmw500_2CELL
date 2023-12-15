import xlrd
import xlwt
import serial
import logging
import time
import re
from xlutils.copy import copy
import time
import matplotlib.pyplot as plt
import numpy as np
import xlrd
from threading import Thread
from datetime import datetime
import os
import random

config = {}
config_pcore = {}
config_acore = {}
def get_config():
    global config, config_pcore
    file_list = [file for file in os.listdir('.') if 'config' in file and file.endswith('.txt')]
    print("请选择配置文件：")
    for i in range(len(file_list)):
        print(str(i) + '、' + file_list[i])
    try:
        n = int(input('请输入测试第几个配置文件：'))
    except:
        n = 0
    print("配置文件名：" + file_list[n])
    # my_log.debug("配置文件名：" + file_list[n])
    with open(file_list[n], 'r', encoding='UTF-8') as file:
        # 逐行读取文件内容
        for line in file:
            # 忽略空行和注释行
            # print(line)
            if '#' not in line:
                if line.strip() == '' or line.strip().startswith('#'):
                    continue
                # 分割每一行的键值对
                if '**' in line.strip():
                    # 分割每一行的键值对
                    key, value = line.strip().split('**')
                    # 添加到全局变量字典中
                    config_pcore[key.strip()] = value.strip()
                elif '==' in line.strip():
                    key, value = line.strip().split('==')
                    # 添加到全局变量字典中
                    config[key.strip()] = value.strip()
                else:
                    if '*' in line.strip():
                        # 分割每一行的键值对
                        key, value = line.strip().split('*')
                        # 添加到全局变量字典中
                        config_acore[key.strip()] = value.strip()
                    else:
                        print("配置文件错误：" + str(line))


get_config()
for key in config_pcore:
    config_pcore[key] = int(config_pcore[key])

for key in config_acore:
    config_acore[key] = int(config_acore[key])

# CT_version = input("电信认证版本选择‘1’，非电信认证版本‘0’：")
# random_power = input("是否做随机断电测试，“1”测试，其他不测试：")
random_power = config['random_power']
if random_power == '1':
    # com2 = input("请输入掉电控制端口号如：com3：")
    com2 = config['random_power']
    ser_control = serial.Serial(port=com2, baudrate=9600, timeout=1)
    print("与下面的掉电模式有冲突，请不要选择‘2’‘3’选项，可选“1”“4”：")
# power = input("cfun0_1-reboot选择’1‘,cfun0_1-断电选择’2‘，CFUN0_断电选择‘3’, CFUN0_reboot选择‘4’：")
power = config['power']

# drawline = input("输入“1”展示曲线，其他不展示：")
drawline = config['drawline']

# psm_mode = input("'1'测试PSM模式，’0‘测试非PSM模式:")
psm_mode = config['psm_mode']

if psm_mode == '1':
    print("建议设置AT+CPSMS=0再测试")
if power == '3' or power == '2':
    # com0 = input("请输入掉电控制端口号如：com3：")
    com0 = config['com0']
    ser_control = serial.Serial(port=com0, baudrate=9600, timeout=1)
    power_port = ser_control.port
# com1 = input("请输入AT串口端口号如：com3：")
com1 =config['com1']
if config['a_p_log'] != '2':
    ser = serial.Serial(port=com1, baudrate=9600, timeout=1)
    test_port = ser.port
    print("AT端口是否打开" + str(ser.isOpen()))
else:
    test_port = 'log'
# 是否自动打印Acore和Pcore log
# a_p_log = input("输入’1‘自动打印log，其他不打印: ")
a_p_log = config['a_p_log']
if config['a_p_log'] != '0':
    # 打开Acore log
    com2 = config['com2']
    # com2 = input("请输入Acore串口端口号如：com3：")
    ser_acore = serial.Serial(port=com2, baudrate=115200, timeout=5)
    # 打开Pcore log
    com3 = config['com3']
    # com3 = input("请输入Pcore串口端口号如：com3：")
    ser_pcore = serial.Serial(port=com3, baudrate=4800000, timeout=5)
    # ser.write(bytes("at+plog=1" + "\r\n", encoding="utf-8"))

# str_CTwing = int(input("是否测试云平台，1 代表测试，其他数字代表不测试："))
# 云平台只有reboot或者断电情况能够自动注册
str_CTwing = int(config['str_CTwing'])
request_data = 0
reject = 0
if psm_mode != '1':
    # at_CSINFO = input("'1'发送AT+LUESTATS=CSINFO，其他不发送:")
    at_CSINFO = config['at_CSINFO']
    if str_CTwing == 1:
        # if power == '1' or power == '2':
        #     auto_CTwing = '0'
        # else:
        #     # auto_CTwing = input("‘1’代表自动注册，其他代表手动注册:")
        auto_CTwing = config['auto_CTwing']
    # ping = int(input("请选择是否ping业务测试，数字 1 代表测试，其他数字
    # 代表不测试："))
    ping = int(config['ping'])
    if ping == 1:
        ping = 1
        # ping_size = input("请输入ping包大小byte：")
        # ping_size = config['ping_size']
    else:
        ping = 0
    # init_test = int(input("是否初始化，’1‘代表初始化，其他代表不初始化："))
    init_test = int(config['init_test'])

    # get_net_data_num = input("是否收集网络数据，1 代表收集，其他代表不收集：")
    get_net_data_num = '0'
    # at_lsdata = input("是否发送AT+LSDATA=0，1 代表发送，其他代表不发送：")
    at_lsdata = config['at_lsdata']

else:
    auto_CTwing = config['auto_CTwing']
    # sleep_time_set = input("请输入休眠设置时间'20','30','60'(实际情况还需要SIM和网络支持)：")
    sleep_time_set = config['sleep_time_set']
    if sleep_time_set == '60':
        str_PSM_time = 'at+cpsms=1,,,,"00100001"'
    elif sleep_time_set == '30':
        str_PSM_time = 'at+cpsms=1,,,,"00001111"'
    else:
        str_PSM_time = 'at+cpsms=1,,,,"00001010"'
    # init_test = int(input("是否初始化，’1‘代表初始化，其他代表不初始化："))
    init_test = int(config['init_test'])
    get_net_data_num = '0'
    at_lsdata = '0'
    ping = 0

# if init_test == 1:
#     init_times = int(input("请输入初始化测试次数："))
# times = input("请输入需要测试的总次数：")

# sleep_time = int(input("请输入每轮测试结束后等待的时间： "))
sleep_time = 0

CTwing_case_num = 0

if str_CTwing == 1:
    # flash = int(input("写入flash输入’0‘，写入sram输入‘1’："))
    flash = 1
    # print('1、透传---明文---注册')
    # print('2、透传---DTLS---注册')
    # print('3、非透传---明文---JSON---注册')
    # print('4、非透传---DTLS---JSON---注册')
    # print('5、透传---明文 IPV4---发送数据')
    # print('6、透传---明文 IPV6---发送数据')
    # print('7、透传---明文 IPV4V6---发送数据')
    # print('8、DTLS---发送数据')
    # print('9、JSON---发送数据')
    # print('10、DTLS---JSON---发送数据"')
    # CTwing_case_num = int(input("请输入云平台测试的类型1,2,3,4,5,6,7,8,9,10: "))
    CTwing_case_num = int(config["CTwing_case_num"])
    CTwing = 1
else:
    CTwing = 0

# ser.flushInput() # 清除串口缓存

# 定义AT命令
str_MR = "AT+CGMR"  # 查看版本信息
str_ADDR = "AT+CGPADDR"  # 获取IP地址
# str_LEARFCN = "AT+LEARFCN=2"  # 清除默认写入频点
str_FLASHZERO = "AT+FLASHZERO=0"  # 清除缓存频点
str_lver = "AT+LVER"  # 查询编译版本信息
# str_BAND5 = "AT+LBAND=5" # 设置band5#
str_CIMI = "AT+CIMI"
str_LCSEARFCN = "AT+LCSEARFCN"  # 清除指定频点和PCI
# str_FLASHZER2 = "AT+FLASHZERO=2"  # 清除Flash
str_REBOOT = "AT+REBOOT"  # 重启UE
str_LUESTATS = "AT+LUESTATS=RADIO"  # 获取小区信息Radio
str_PING = "AT+LPING=221.229.214.202,10000,60,20"  # 发起PING，长度60，20次
str_CPSMS_1 = "AT+CPSMS=1"  # 进入PSM模式
str_CPSMS_0 = "AT+CPSMS=0"  # 禁用PSM模式
str_CPSMS = "AT+CPSMS?"  # 查询PSM模式状态
str_LVER = "AT+LVER"  # 获取编译版本信息
str_AT = 'AT'
str_CFUN0 = "AT+CFUN=0"
str_CFUN1 = "AT+CFUN=1"
str_disable_ldo = "disable_ldo"
str_enable_ldo = "enable_ldo"
str_CELL = 'AT+LUESTATS=CELL'
str_CSINFO = 'AT+LUESTATS=CSINFO'
str_WATCHDOG_0 = 'AT+WATCHDOG=0'
str_MONITOR_1 = "AT+MONITOR=1"
str_WATCHDOG_1 = 'AT+WATCHDOG=1'
str_MONITOR_0 = "AT+MONITOR=0"
str_LRFTCF = "AT+LRFTCF=1"  # LSCLK=0,CPSMS=1,PSDEBUG=1,WATCHDOG=1,MONITOR=1,LSEFREG=0,band=5,IPV4,APN=CTNB
str_LRFTCF_1 = "AT+LRFTCF=1,0,1"  # 恢复出厂设置PSM模式关闭
str_LRFTCF_2 = "AT+LRFTCF=1,1,1"  # 恢复出厂设置PSM模式打开
str_LCSEARFCN = "AT+LCSEARFCN"  # 清除指定频点和PCI
str_LEARFCN_2057 = "AT+LEARFCN=1,2507"  # 锁定频点为2507
str_LEARFCN_2059 = "AT+LEARFCN=1,2509"  # 锁定频点为2509
str_LSERV = "AT+LSERV=221.229.214.202,5683"  # 设置服务器
str_LCTM2MREG = "AT+LCTM2MREG"  # 注册
str_LCTM2MDEREG = "AT+LCTM2MDEREG"  # 去注册
str_LCTM2MSEND = "AT+LCTM2MSEND=112233"  # 发送数据给云平台
str_LCTM2MINT = "AT+LCTM2MINIT=869951046023283"
str_LSERV_MQTT = "AT+LSERV=180.106.148.146,1883"
str_LMQTTCON = 'AT+LMQTTCON="1523357612345","6nJTv44lckuPgrNsBntrpnKIVG1KYKq-E8roliHxaJ0"'
str_NBSTATS = 'AT+NBSTATS'  # 查询监控网络信息
str_LWIPDATA = 'AT+LWIPDATA'  # 查询监控云平台注册信息
str_LSDATA = 'at+lsdata=0'  # cfun=0回ok后不保存历史频点


# 定义全局变量
# j = 0
AT_die = 0
AT_death = 0
long_time = 0
CELLID_error = 0
CTwing_reg = 0
CTwing_data = 0
CTwing_sub = 0
expected_cellid = '504'
psm_get_net_time = 0
psm_active = 0
rsrp_old = 0
# 定义死机复位种类
RST_0 = 0
RST_1 = 0
RST_2 = 0
RST_3 = 0
RST_4 = 0
RST_5 = 0
RST_6 = 0
# 定义入网开始时间
t1 = 0
restart_times = 0
restart_reboot = 0
restart_power = 0
restart_psm = 0
# 定义打印excel list()
list = []
list1 = []
list2 = []
list_result = []
# PSM模式下测试数据
list3 = []

# 定义ping业务丢包
ping_1 = 0
ping_2 = 0
ping_all = 0
ping_a = 0
# 统计网络掉线次数
CEREG_0 = 0

# 定义百分比
t_rate = 0
at_rate = 0
cell_rate = 0
CT_rate = 0
CT_data_rate = 0
# 网络注销时间
detach_time = '0'
CTwing_deteach = 0
imei = "Fail"
# psm模式业务结束时间
task_end_time = 0
at_time = 0
drawlinepsm = 0
draw_line_time = 0
rsrp_change_times = 0
dic_sim = 0

flash_SPI = 0

flash_change = 0

ping_total_all = 0


#  获取当前时间
time_excel = time.strftime('%m-%d-%H-%M-%S', time.localtime(time.time()))
# 创建文件夹
path_current = os.getcwd()
path = os.getcwd() + '\\' + time_excel + test_port
# os.path.exists(path)
os.mkdir(path)
# print(os.path.exists(path))

# 打印log
my_log = logging.getLogger()
my_log.setLevel(level=logging.DEBUG)

# 去除maplat乱码打印
logging.getLogger('matplotlib').setLevel(logging.WARNING)
my_log.addFilter(lambda record: 'STREAM' not in record.getMessage())

filehandle01 = logging.FileHandler(path + '\\' + time_excel + "my_log" + test_port + ".txt", 'w', 'utf-8')
filehandle02 = logging.FileHandler(path + '\\' + time_excel + "my_report" + test_port + ".txt", 'w', 'utf-8')

filehandle02.setLevel(logging.ERROR)
# formatter01 = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
formatter01 = logging.Formatter('%(asctime)s -%(levelname)s - %(message)s')
formatter02 = logging.Formatter('%(asctime)s -%(message)s')

filehandle01.setFormatter(formatter01)
filehandle02.setFormatter(formatter02)

my_log.addHandler(filehandle01)
my_log.addHandler(filehandle02)


# def init():
#     global t1, restart_reboot, list, psm_mode, task_end_time, at_first_time
#     print("初始化........")
#     my_log.debug("初始化........")
#     send_AT0('AT')
#     time.sleep(0.1)
#     if psm_mode == '1':
#         send_AT0(str_LRFTCF_2)
#         time.sleep(0.2)
#         send_AT0(str_LRFTCF_2)
#         time.sleep(1)
#         send_AT0("AT+PLOG=1")
#         time.sleep(1)
#         send_AT0('at+lselfreg=0')
#         time.sleep(1)
#         t1 = time.time()
#         at_first_time = t1
#         send_AT0("AT+REBOOT")
#         send_AT0(str_PSM_time)
#         t = get_net_time()
#         my_log.debug("初始化第一次入网时间：" + str(t))
#         print("初始化第一次入网时间：" + str(t))
#         print("下面等待进入休眠...")
#         my_log.debug("下面等待进入休眠...")
#         task_end_time = time.time()
#         psm_test()
#         return list3
#
#     else:
#         send_AT0(str_LRFTCF_1)
#         time.sleep(0.2)
#         send_AT0(str_LRFTCF_1)
#         time.sleep(1)
#         send_AT0("AT+PLOG=1")
#         time.sleep(1)
#         send_AT0('at+lselfreg=0')
#         time.sleep(1)
#         t1 = time.time()
#         send_AT0("AT+REBOOT")
#         restart_reboot += 1
#         # 初始化reboot 之后进行一轮测试
#         t = get_net_time()
#         if t == -1:
#             list = ['0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '-1', '0', '0', '0', '25', '0', '0', '0', '0', '0','0']
#             list1.append(list[0])
#             list1.append(list[17])
#             list2.append(list[3])
#             # 获取网络注册信息
#             get_net_data()
#             re_start()
#             return list
#             # list.append(t)
#         list.append(t)
#         list.append(detach_time)
#         # 云平台测试
#         CTWing_case_run(CTwing, CTwing_case_num)
#         # 获取小区信息
#         get_cell()
#         get_cell_csinfo()
#
#         # 做Ping业务
#         do_ping(ping)
#         my_log.debug(list)
#         list1.append(list[0])
#         list1.append(list[17])
#         list2.append(list[3])
#         get_net_data()
#         return list

def wait_str(value):
    data = getdata_line()
    i = 0
    while value not in data:
        # time.sleep(1)
        data = getdata_line()
        i = i + 1
        if i > 10:
            return "fail"
    else:
        return 'pass'


def init():
    global t1, restart_reboot, list, psm_mode, task_end_time, at_first_time, flash, CTwing, imei
    def cfun0():
        getdata_0()
        send_AT0("AT+CFUN=0")
        result_at = wait_str("OK")
        print("CFUN0:" + str(result_at))
        my_log.debug("CFUN0:" + str(result_at))
        return result_at
    result_at = cfun0()
    while result_at == "fail":
        print("cfun0失败，reboot重新发送")
        send_AT0('AT+REBOOT')
        restart_reboot += 1
        getdata_sleep(5)
        result_at = cfun0()
    getdata_sleep(3)
    print("初始化........")
    my_log.debug("初始化........")
    getdata_0()
    send_AT0(str_LRFTCF)
    str_LRFTCF_1 = getdata_0()
    while 'Alert!' not in str_LRFTCF_1:
        str_LRFTCF_1 = getdata_0()
        time.sleep(1)
        send_AT0(str_LRFTCF)
    else:
        send_AT0(str_LRFTCF)
        str_LRFTCF_2 = getdata_0()
        i = 0
        while 'restore to factory configuration' not in str_LRFTCF_2:
            str_LRFTCF_2 = getdata_0()
            time.sleep(0.5)
            i += 1
            if i == 10:
                print("初始化失败，建议复位后重新开始测试")
        else:
            getdata_0()
            if config['plog'] == '1':
                send_AT0("AT+PLOG=1")
            else:
                send_AT0("AT+PLOG=0")
            getdata_sleep(1)
            if config['learfcn0'] != '0':
                # 锁定频点
                send_AT0(config['learfcn0'])
                print("锁频点" + config['learfcn0'])
            else:
                # 清除或者写回候选频点
                send_AT0(config['learfcn'])
                print("候选频点" + config['learfcn'])
            getdata_sleep(1)
            # if config['flash_p'] == '1':
            #     print('开启flash保护')
            #     my_log.debug('开启flash保护')
            #     if config['a_p_log'] == '1':
            #         ser_acore.write(bytes("flash p" + "\r\n", encoding="utf-8"))
            #     else:
            #         print('未开启acore')
            #         my_log.debug('未开启acore')
            # else:
            #     print("未开启flash保护")
            #     my_log.debug("未开启flash保护")


            if CTwing == 1:
                print("配置云平台参数" + str(CTwing_case_num))
                CTWing_case_config(CTwing_case_num, imei)
                print('云平台自动注册:' + auto_CTwing)
                if auto_CTwing == "1":
                    send_AT('at+lctregen=1')
                else:
                    send_AT('at+lctregen=0')
                if flash == 0:
                    send_AT0(str_REBOOT)
                    restart_reboot += 1
            else:
                send_AT('at+lctregen=0')
            # 发送CFUN1，写入PSM模式
            send_AT0("AT+CFUN=1")
            wait_str('OK')
            t1 = time.time()
            if psm_mode == '1':
                getdata_0()
                if a_p_log == '1':
                    tt2 = Thread(target=read_pcore)
                    tt2.start()
                send_AT0(str_PSM_time)
                wait_str('OK')
                send_AT0("AT+REBOOT")
                print("初始化reboot后等待入网....")
                getdata_0()
                restart_reboot += 1
                at_first_time = t1
                t = get_net_time()
                my_log.debug("初始化第一次入网时间：" + str(t))
                print("初始化第一次入网时间：" + str(t))
                set_psm_mode()
                check_lsclk()
                print("下面等待进入休眠...")
                my_log.debug("下面等待进入休眠...")
                task_end_time = time.time()
                psm_test()
                getdata_0()
                if config['learfcn0'] != '0':
                    # 锁定频点
                    send_AT0(config['learfcn0'])
                    print("锁频点" + config['learfcn0'])
                else:
                    print("写回候选频点")
                    send_AT0(config['learfcn1'])

                getdata_0()
                return list3

            else:
                # 初始化reboot 之后进行一轮测试
                getdata_0()
                if a_p_log == '1':
                    tt2 = Thread(target=read_pcore)
                    tt2.start()
                send_AT0("AT+CPSMS=0")
                wait_str('OK')
                getdata_0()
                t1 = time.time()
                send_AT0("AT+REBOOT")
                restart_reboot += 1
                getdata_0()
                t = get_net_time()
                set_psm_mode()
                # check_lsclk()
                if t == -1:
                    list = ['-1', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '9999', '0', '500']
                    list1.append(list[0])
                    list1.append(list[17])
                    list2.append(list[3])
                    # 获取网络注册信息
                    get_net_data()
                    send_AT0(config['learfcn1'])
                    time.sleep(1)
                    re_start()
                    return list
                    # list.append(t)
                my_log.debug("初始化第一次入网时间：" + str(t))
                print("初始化第一次入网时间：" + str(t))
                list.append(t)
                list.append(detach_time)
                # 云平台测试
                CTWing_case_run(CTwing, CTwing_case_num)
                # 获取小区信息
                get_cell()
                get_cell_csinfo()
                # 做Ping业务
                do_ping(ping)
                my_log.debug(list)
                list1.append(list[0])
                list1.append(list[17])
                list2.append(list[3])
                get_net_data()
                send_AT0(config['learfcn1'])
                getdata_0()
                return list


def get_imei():
    # ser.write(bytes("AT+CGSN" + "\r\n", encoding="utf-8"))
    send_AT0("AT+CGSN")
    # i = 0
    time.sleep(1)
    data = getdata_0()
    # data = str(ser.read_all())
    if "OK" and '8622' not in data:
        print("输入命令：" + "AT+CGSN" + "失败")
        return "Fail"
    else:
        data_imei = re.findall(r'\d+', data.strip())
        print(data_imei[0])
        return data_imei[0]


def random_power1():
    global restart_power, psm_mode
    # num = random.randint(0,100)
    while(1):
        random.seed(int(datetime.now().timestamp()))
        num = random.randint(0, 100)
        # print(num)
        # print(num)
        if psm_mode == '1':
            time.sleep(5)
        else:
            time.sleep(2)
        if num < 5:
            restart_power += 1
            getdata_0()
            ser_control.write((bytes(str_disable_ldo + "\r\n", encoding="utf-8")))
            print("随机断电")
            my_log.debug("随机断电")
            time.sleep(2)
            ser_control.write((bytes(str_enable_ldo + "\r\n", encoding="utf-8")))
            print("随机上电")
            my_log.debug("随机上电")


def flash_u():
    ser_acore.write(bytes("flash u" + "\r\n", encoding="utf-8"))
    print("关闭flash保护")
    my_log.debug("关闭flash保护")


def read_acore():
    global request_data, config_acore, flash_change, flash_SPI, RST_3
    print("启动Acore线程")
    name = path + '\\' + time_excel + ser_acore.port + 'Acorelog' + '.txt'
    A_P_RAM_SHARE = 0
    acore_rrc_req_time_0 = time.time()
    acore_rrc_time_0 = time.time()
    acore_auth_time_0 = time.time()
    acore_secu_time_0 = time.time()
    acore_esm_time_0 = time.time()
    attach_complete_time_0 = time.time()
    with open(name, 'a', encoding='ASCII') as f:
        while(ser_acore.isOpen()):
            try:
                data_acore = ser_acore.readline()
            except Exception as e:
                print(e)
                my_log.debug(e)
                data_acore = ser_acore.read_all()
            try:
                data2 = data_acore.decode('ASCII').strip() + "\n"
                data1 = "[" + str(datetime.now().strftime('%m-%d %H:%M:%S.%f')) + ']' + ' '
                data = str(data1 + data2)
                if 'ACORE_PCORE_RAM_SHARE' in data2:
                    A_P_RAM_SHARE += 1
                    my_log.debug(data2)
                # if A_P_RAM_SHARE >= 1:
                for key in config_acore:
                    if key in data2:
                        config_acore[key] += 1

                if re.search(r'S.+G.+.K.+S', data2) != None or re.search(r'G....K....D', data2) != None:
                    my_log.debug("-----------------------------------------flash数据被篡改---------------------------------------")
                    my_log.debug(data2)
                    flash_change += 1
                    # print(data2)
                # if re.search(r'SPITFE', data2) != None or re.search(r'SPIRFNE', data2) != None or re.search(r'SPIRXFLR', data2) != None or re.search(r'SPIBUSY', data2) != None:
                if re.search(r'acore will reboot PCORE AuthFail', data2) != None:
                    my_log.debug("-----------------------------------------鉴权失败---------------------------------------")
                    my_log.debug(data2)
                    RST_3 += 1
                    # flash_SPI += 1
                    # print(data2)
                if 'MAILBOX_RRC_REQ_SUCCESS 0' in data2:
                    acore_rrc_req_time_0 = time.time()
                if 'MAILBOX_RRC_REQ_SUCCESS 1' in data2:
                    rrc_req_time = round((time.time() - acore_rrc_req_time_0), 2)
                    my_log.debug("MAILBOX_RRC_REQ_SUCCESS_0_1时间：" + str(rrc_req_time))
                    # print("MAILBOX_RRC_REQ_SUCCESS_0_1时间：" + str(rrc_req_time))

                if 'MAILBOX_RRC_CONN_COMPLETE_SUCCESS 0' in data2:
                    acore_rrc_time_0 = time.time()
                if 'MAILBOX_RRC_CONN_COMPLETE_SUCCESS 1' in data2:
                    rrc_complete_time = time.time()
                    rrc_time = round((rrc_complete_time - acore_rrc_time_0), 2)
                    my_log.debug("RRC_CONN_COMPLETE_SUCCESS_0_1时间：" + str(rrc_time))
                    # print("RRC_CONN_COMPLETE_SUCCESS_0_1时间：" + str(rrc_time))

                if 'MAILBOX_AUTH_RESP_SUCCESS 0' in data2:
                    acore_auth_time_0 = time.time()
                if 'MAILBOX_AUTH_RESP_SUCCESS 1' in data2:
                    auth_time = round((time.time() - acore_auth_time_0), 2)
                    my_log.debug("MAILBOX_AUTH_RESP_SUCCESS_0_1时间：" + str(auth_time))
                    # print("MAILBOX_AUTH_RESP_SUCCESS_0_1时间：" + str(auth_time))

                if 'MAILBOX_SECU_MODE_COMPLETE_SUCCESS 0' in data2:
                    acore_secu_time_0 = time.time()
                if 'MAILBOX_SECU_MODE_COMPLETE_SUCCESS 1' in data2:
                    secu_time = round((time.time() - acore_secu_time_0), 2)
                    my_log.debug("MAILBOX_SECU_MODE_COMPLETE_SUCCESS_0_1时间：" + str(secu_time))
                    # print("MAILBOX_SECU_MODE_COMPLETE_SUCCESS_0_1时间：" + str(secu_time))

                if 'MAILBOX_ESM_INFO_RESP_SUCCESS 0' in data2:
                    acore_esm_time_0 = time.time()
                if 'MAILBOX_ESM_INFO_RESP_SUCCESS 1' in data2:
                    esm_time = round((time.time() - acore_esm_time_0), 2)
                    my_log.debug("MAILBOX_ESM_INFO_RESP_SUCCESS_0_1时间：" + str(esm_time))
                    # print("MAILBOX_ESM_INFO_RESP_SUCCESS_0_1时间：" + str(esm_time))

                if 'MAILBOX_ATTACH_COMPLETE_SUCCESS 0' in data2:
                    attach_complete_time_0 = time.time()
                    rrc_attach_complet_time = round((attach_complete_time_0 - rrc_complete_time), 2)
                    my_log.debug("MAILBOX_RRC_1_ATTACH_COMPLETE_SUCCESS_0时间：" + str(rrc_attach_complet_time))
                    # print("MAILBOX_RRC_1_ATTACH_COMPLETE_SUCCESS_0时间：" + str(rrc_attach_complet_time))
                if 'MAILBOX_ATTACH_COMPLETE_SUCCESS 1' in data2:
                    attach_complete_time = round((time.time() - attach_complete_time_0), 2)
                    my_log.debug("MAILBOX_ATTACH_COMPLETE_SUCCESS_0_1时间：" + str(attach_complete_time))
                    # print("MAILBOX_ATTACH_COMPLETE_SUCCESS_0_1时间：" + str(attach_complete_time))


                f.write(data)
                f.flush()

            except Exception as e:
                # print(e.__str__())
                f.write(e.__str__())
                my_log.debug(data_acore)


def read_pcore():
    global request_data, reject, SIMST_0, dic_sim, config_pcore
    print("启动Pcore线程")
    NBsys = 0
    name = path + '\\' + time_excel + ser_pcore.port + 'Pcorelog' + '.txt'
    with open(name, 'a', encoding='ASCII') as f:
        while(ser_pcore.isOpen()):
            try:
                data_pcore = ser_pcore.readline()
            except Exception as e:
                print(e)
                my_log.debug(e)
                data_pcore = ser_pcore.read_all()
            # print(data_pcore)
            try:
                data2 = data_pcore.decode('ASCII').strip() + "\n"
                data1 = "[" + str(datetime.now().strftime('%m-%d %H:%M:%S.%f')) + ']' + ' '
                data = str(data1 + data2)
                if a_p_log != '2':
                    if 'NBsys' in data2:
                        NBsys += 1
                    if NBsys >= 1:
                        for key in config_pcore:
                            if key in data2:
                                config_pcore[key] += 1
                        if re.search(r'Paddy.+(attention|error|timeout)', data2) != None:
                            dic_sim += 1
                            my_log.debug(data2)
                f.write(data)
                f.flush()
            except Exception as e:
                # print(e.__str__())
                # data = "-----------------------打印乱码异常了-------------------------"
                my_log.debug(str(data_pcore))
                f.write(e.__str__())


def set_psm_mode_0():
    global restart_reboot
    getdata_0()
    send_AT0("AT+CPSMS?")
    data = getdata_0()
    i = 0
    if 'LETSWIN' in data:
        send_AT0("AT+CPSMS?")
        data = getdata_0()
    while 'CPSMS: 0' not in data:
        data = getdata_line()
        i = i + 1
        if i % 20 == 5:
            send_AT0("AT+CPSMS?")
        if i % 20 == 15:
            print("PSM模式0设置不成功,reboot重新配置")
            my_log.debug("PSM模式0设置不成功,reboot重新配置")
            send_AT0("AT+CPSMS=0")
            getdata_sleep(0.5)
            send_AT0('AT')
            send_AT0("AT+REBOOT")
            restart_reboot += 1
            getdata_sleep(8)
    else:
        print("PSM模式0设置成功")


def set_psm_mode():
    global t1
    global restart_reboot, psm_mode
    if psm_mode != '1':
        getdata_0()
        send_AT0("AT+CPSMS?")
        data = getdata_line()
        i = 0
        while 'CPSMS: 0' not in data:
            data = getdata_line()
            i = i + 1
            if i % 20 == 5:
                send_AT0("AT+CPSMS?")
            if i % 20 == 15:
                print("PSM模式0设置不成功,reboot重新配置")
                send_AT0("AT+CPSMS=0")
                getdata_sleep(1)
                t1 = time.time()
                send_AT0("AT+REBOOT")
                getdata_sleep(2)
                t = get_net_time()
                my_log.debug("psm模式设置失败reboot后入网时间：" + str(t))
        else:
            print("PSM模式0设置成功")
    else:
        if sleep_time_set == '60':
            PSM_march = 'CPSMS: 1,,,"00001101""00100001"'
        elif sleep_time_set == '30':
            PSM_march = 'CPSMS: 1,,,"00001101""00001111"'
        else:
            PSM_march = 'CPSMS: 1,,,"00001101""00001010"'
        getdata_0()
        # send_AT0(str_PSM_time)
        # time.sleep(1)
        send_AT0("AT+CPSMS?")
        data = getdata_line()
        i = 0
        while PSM_march not in data:
            if 'CPSMS: 1' in data and i > 60:
                input("请确认是不是PSM卡，继续按任意键：")
                break
            data = getdata_line()
            i = i + 1
            if i % 10 == 5:
                send_AT0("AT+CPSMS?")
            if i % 30 == 9:
                print("PSM模式设置不成功,重新配置")
                send_AT0(str_PSM_time)
            if i % 30 == 13:
                print("PSM模式设置不成功, 现在reboot")
                send_AT0("AT")
                send_AT0("AT+REBOOT")
                restart_reboot += 1
                getdata_sleep(10)

        else:
            print("PSM模式设置成功")


def check_lsclk():
    getdata_0()
    send_AT0("AT+LSCLK?")
    if psm_mode == '1':
        data = getdata_line()
        i = 0
        while "LSCLK: 2\\r\\n" not in data:
            data = getdata_line()
            i = i + 1
            if i % 10 == 5:
                send_AT0("AT+LSCLK?")
            if i%10 == 9:
                print("AT+LSCLK查询结果，不等于2")
                send_AT0("AT+LSCLK=2")
            if i == 30:
                init("如需要继续，任意键盘继续测试")
                break
        else:
            print("AT+LSCLK查询结果为2")
    else:
        data = getdata_line()
        i = 0
        while 'LSCLK: 0' not in data:
            data = getdata_line()
            i = i + 1
            if i % 10 == 5:
                send_AT0("AT+LSCLK?")
            if i % 10 == 9:
                print("AT+LSCLK查询结果一次，不等于0")
                send_AT0("AT+LSCLK=0")
        else:
            print("AT+LSCLK查询结果为0")


def psm_mode_1():
    global restart_reboot
    send_AT(str_AT)
    time.sleep(0.2)
    send_AT(str_CPSMS_1)
    time.sleep(1)
    my_log.debug(getdata())
    # send_AT3(str_CPSMS)
    ser.write(bytes(str_CPSMS + "\r\n", encoding="utf-8"))
    my_log.debug("发送AT命令: " + str_CPSMS)
    time.sleep(0.5)
    read_data = getdata()
    i = 0
    while "+CPSMS: 1" not in read_data:
        if i < 3:
            time.sleep(1)
            i = i + 1
            send_AT(str_CPSMS_1)
            time.sleep(1)
            ser.write(bytes(str_CPSMS_1 + "\r\n", encoding="utf-8"))
            my_log.debug("发送AT命令: " + str_CPSMS)
            read_data = getdata()
            my_log.debug(read_data)
            # print(read_data)
        else:
            # send_AT(str_REBOOT)
            # time.sleep(1)
            # restart_reboot = restart_reboot + 1
            # getdata()
            return "error"
    else:
        # send_AT(str_REBOOT)
        # time.sleep(1)
        # restart_reboot = restart_reboot + 1
        # getdata()
        return "Yes"


# 解除PSM模式，成功返回“Yes”，不成功返回“error”
def psm_mode_0():
    global restart_reboot, t1
    send_AT0(str_AT)
    time.sleep(2)
    send_AT(str_CPSMS_0)
    time.sleep(1)
    send_AT0(str_CPSMS)
    read_data = getdata()
    i = 0
    while "+CPSMS: 0" not in read_data:
        if i < 3:
            time.sleep(1)
            i = i + 1
            send_AT0(str_CPSMS_0)
            time.sleep(0.3)
            ser.write(bytes(str_CPSMS + "\r\n", encoding="utf-8"))
            read_data = getdata()
        else:
            send_AT0(str_REBOOT)
            time.sleep(1)
            restart_reboot = restart_reboot + 1
            getdata()
            return "error"
    else:
        send_AT0(str_REBOOT)
        t1 = time.time()
        my_log.debug("psm设置后开始计算入网时间")
        time.sleep(1)
        restart_reboot = restart_reboot + 1
        getdata()
        return "Yes"


def psm_test():
    global restart_psm, psm_active, psm_get_net_time, t1, task_end_time, list3, str_CTwing
    # 等待休眠,返回int入网时间
    wait_sleep_start = time.time()
    sleep1 = wait_sleep()
    if sleep1[0] == 'PSM超时120s':
        list3 = [-1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]    # 发送AT，等待入网，返回int入网时间
        re_start()
        return list3
    else:
        sleep1_float = float(sleep1[0])
    # 获取入网时间，返回”入网失败“
    net_time = at_net(sleep1_float)
    if net_time == '入网失败':
        # print(list3)
        while len(list3) < 13:
            list3.append('0')
        re_start()
        return list3
    # 获取cell信息，无返回值
    get_psm_cell()
    # 注册登录，注册云平台，返回t4, '订阅成功'
    if str_CTwing == 1:
        if config['psm_auto_send_data'] == '1':
            ctwing_result = ['0', '订阅成功']
            list3.append("-1")
            list3.append("-1")
        else:
            ctwing_result = wait_CTwing_reg()
        if ctwing_result[1] == '订阅成功':
            # 发送云平台数据，返回发送数据时间
            send_data_result = send_message(
                data='"{"id":"11:22:33:44:55:66","rssi":"[{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93}]""gps":"113.58761666,37.86236085","time":"1681182061"}"')
            # list3.append(send_data_result)
            # 注销云平台
            if send_data_result == "发送数据30s超时":
                # dereg_result = DEREG_1()
                # 补齐注销云平台数据
                list3.append("-1")
            else:
                # dereg_result = DEREG()
                list3.append("-1")
        else:
            list3.append("0")
            list3.append("-1")
    else:
        list3.append("0")
        list3.append("0")
        list3.append("0")
        list3.append("0")
    # 计算从deteach到deep sleep时间
    # ，业务到sleep时间
    task_end_time = time.time()
    # print("all_time: " + str(task_end_time - wait_sleep_start))

    list3.append(format(task_end_time - sleep1[1], '.2f'))
    my_log.debug("task_end_time: " + str(task_end_time))


def wait_sleep():
    global restart_psm, psm_active, psm_get_net_time, t1, task_end_time, at_time, at_first_time, sleep_time_set,draw_line_time
    my_log.debug("等待进入深睡............")
    print("等待进入深睡：", flush=True, end='')
    # 输入AT防止进入深睡后没发激活
    at_time = time.time()
    getdata_0()
    send_AT0(str_AT)
    data = getdata_line()
    while "DEEP SLEEP" not in data:
        data = getdata_line()
        print('.', end='', flush=True)
        psm_t1 = time.time()
        # 计算从入网到进入休眠的时间
        net_sleep_time = psm_t1 - t1
        net_sleep_time2 = psm_t1 - at_time
        limit_time = int(config['sleep_limit_time'])

        if net_sleep_time > limit_time and net_sleep_time2 > limit_time:
            print("等待深睡超过" + str(limit_time) + "秒")
            my_log.error("等待深睡超过" + str(limit_time) + "秒")
            my_log.error("Fail")
            list3.append("-1")
            psm_active += 1
            return "PSM超时120s", at_time
    else:
        # 业务结束到获取DEEP SLEEP时间
        # psm_sleep_time_1 = time.time()
        task_time = format((time.time() - task_end_time), '.2f')
        # -float(draw_line_time)
        # my_log.debug("休眠时间：000000000")
        all_time = format((time.time() - at_first_time), '.2f')
        list3.append(all_time)
        at_first_time = at_time
        my_log.debug('at_first_time: ' + str(at_first_time))
        # print("所有业务的时间：" + str(all_time))
        list3.append(task_time)
        psm_sleep_time = format((time.time() - at_time), '.2f')
        my_log.debug("入网后到进入休眠的时间: " + psm_sleep_time)
        return psm_sleep_time, at_time


def at_net(sleep1_float):
    global restart_psm, psm_get_net_time, sleep_time_set, send_AT_time
    print("进入深睡，下面发送AT命令开始激活等待入网")
    # 进入深睡后休眠1s，防止异常出现
    if 'wake_time' in config:
        time.sleep(int(config['wake_time']))
    else:
        time.sleep(1)
    psm_t2 = time.time()
    getdata_0()
    send_AT_time = time.time()
    my_log.debug('send_AT_time' + str(send_AT_time))
    send_AT0(str_AT)
    restart_psm = restart_psm + 1
    psm_active_data = getdata_line()
    while "PSM ACTIVE" not in psm_active_data:
        psm_active_data = getdata_line()
        ac_time = float(format((time.time() - send_AT_time), '.2f'))
        limit_time = int(config['net_limit_time'])
        if ac_time > limit_time:
            list3.append("-1")
            my_log.debug('等待' + str(limit_time) + 's入网失败')
            print("psm入网失败")
            psm_get_net_time += 1
            return "入网失败"
    else:
        # 从active到注册网络时间，写入excel
        psm_net_time = format((time.time() - send_AT_time), '.2f')
        list3.append(psm_net_time)
        return psm_net_time


def send_AT(command):
    global AT_die
    '''发送AT命令如果成功返回OK，如果不成功返回“Fail”'''
    ser.write(bytes(command + "\r\n", encoding="utf-8"))
    my_log.debug("输入命令" + command)
    time.sleep(0.5)
    # i = 0
    data = getdata_0()
    if "OK" not in data:
        # i = i + 1
        # my_log.debug(data)
        my_log.debug("输入命令：" + command + "失败")
        # data = str(ser.read_all())
        print("输入命令：" + command + "失败")
        AT_die = AT_die + 1
        return "Fail"
    else:
        # my_log.debug(data)
        my_log.debug("输入命令：" + command + "成功")
        return "Pass"


# 每0.7s获取串口数据并且解码
def getdata():
    global RST_0, RST_1, RST_2, RST_3, RST_4, RST_5, RST_6
    count = ser.inWaiting()
    # print("缓存区数据大小" + str(count))
    if count > 0:
        time.sleep(0.7)
        data = ser.read_all().decode('utf-8')
        # print(data)
        read = data.split('\n')
        # print(read)
        read_data = '\n'.join(read)
        my_log.debug(read_data)
        if 'RST(0)' in data:
            RST_0 += 1

        if 'RST(1)' in data:
            RST_1 += 1

        if 'RST(2)' in data:
            RST_2 += 1

        if 'RST(3)' in data:
            RST_3 += 1

        if 'RST(4)' in data:
            RST_4 += 1

        if 'RST(5)' in data:
            RST_5 += 1

        if 'RST(6)' in data:
            RST_6 += 1
        # if "DEEP SLEEP" in data:
        #     print("进入了PSM模式，现在关闭")
        #     print(psm_mode_0())
        return read_data
    else:
        return "数据为空"


def getdata_sleep(times):
    global RST_0, RST_1, RST_2, RST_3, RST_4, RST_5, RST_6
    t_sleep1 = time.time()
    t = 0
    while t < times:
        data = str(ser.readline())
        my_log.debug(data + '----------------------------------------------------------------------')

        if data is not None and 'RST(0)' in data:
            RST_0 += 1

        if data is not None and 'RST(1)' in data:
            RST_1 += 1

        if data is not None and 'RST(2)' in data:
            RST_2 += 1

        if data is not None and 'RST(3)' in data:
            RST_3 += 1

        if data is not None and 'RST(4)' in data:
            RST_4 += 1

        if data is not None and 'RST(5)' in data:
            RST_5 += 1

        if data is not None and 'RST(6)' in data:
            RST_6 += 1
        # if "DEEP SLEEP" in data:
        #     print("进入了PSM模式，现在关闭")
        #     print(psm_mode_0())
        t = time.time() - t_sleep1


def getdata_0():
    global RST_0, RST_1, RST_2, RST_3, RST_4, RST_5, RST_6
    data = str(ser.read_all())
    my_log.debug(data)

    if data is not None and 'RST(0)' in data:
        RST_0 += 1

    if data is not None and 'RST(1)' in data:
        RST_1 += 1

    if data is not None and 'RST(2)' in data:
        RST_2 += 1

    if data is not None and 'RST(3)' in data:
        RST_3 += 1

    if data is not None and 'RST(4)' in data:
        RST_4 += 1

    if data is not None and 'RST(5)' in data:
        RST_5 += 1

    if data is not None and 'RST(6)' in data:
        RST_6 += 1
    # if "DEEP SLEEP" in data:
    #     print("进入了PSM模式，现在关闭")
    #     print(psm_mode_0())
    return data


def getdata_0_0():
    # global RST_0, RST_1, RST_2, RST_3, RST_4, RST_5
    data = str(ser.read_all())
    my_log.debug(data)
    return data


def getdata_line():
    global RST_0, RST_1, RST_2, RST_3, RST_4, RST_5, RST_6
    data = str(ser.readline())
    my_log.debug(data)
    if 'RST(0)' in data:
        RST_0 += 1
        # while RST_0 > 0:
        #     pass

    if 'RST(1)' in data:
        RST_1 += 1

    if 'RST(2)' in data:
        RST_2 += 1

    if 'RST(3)' in data:
        RST_3 += 1
        # while RST_3 > 0:
        #     pass

    if 'RST(4)' in data:
        RST_4 += 1

    if 'RST(5)' in data:
        RST_5 += 1

    if 'RST(6)' in data:
        RST_6 += 1

    return data


def get_ip_addr():
    ser.write(bytes(str_ADDR + "\r\n", encoding="utf-8"))
    time.sleep(1)
    gd = getdata()
    # print("获取IP结果：" + gd)
    # print(time.time())
    if 'ADDR' not in gd:
        # print("未获取到IP地址")
        return 0
    else:

        return 1


def send_AT0(command):
    my_log.debug("输入命令" + command)
    ser.write(bytes(command + "\r\n", encoding="utf-8"))
    time.sleep(0.5)
    # print("输入命令" + command)


def get_net_time():
    '''判断是否入网，返回统计时间t'''
    global AT_die, long_time, CELLID_error, t1, AT_death,CT_version
    i = 0
    # getdata_sleep(1)
    net = getdata_line()
    if 'net_timeout' in config:
        net_timeout = int(config['net_timeout'])
    else:
        net_timeout = 120
    print("入网等待时间" + str(net_timeout) + 's....')
    my_log.debug("等待是否入网........：")
    while 'ADDR' not in net:
        # time.sleep(0.5) 行读取设置串口读取超时时间为1s，故不设置等待时间
        net = getdata_line()
        i += 1
        t4 = time.time() - t1
        # print(t4)
        if 20 < t4 < 21:
            send_AT0(str_ADDR)
            time.sleep(0.5)
            net = getdata_line()
            # return -1

        if net_timeout < t4 < net_timeout + 1:
            # if config['CT_version'] != '1':
            #     get_ota_data()
            send_at_result = send_AT(str_ADDR)
            if send_at_result == "Pass":
                my_log.debug(str(net_timeout) + "s不入网后输入" + str_ADDR + " Acore正常打印")
                print(str(net_timeout) + "s不入网后输入" + str_ADDR + " Acore正常打印")
            if send_at_result == "Fail":
                AT_death = AT_death + 1
            # 等待2s，不等待马上重启可能出现Pcorelog没打印
            time.sleep(2)
            long_time = long_time + 1

            send_AT0(str_ADDR)
            time.sleep(0.5)
            net = getdata_line()

        if t4 > net_timeout + 1:
            print("入网时间超过" + str(net_timeout) + "s,现在restart：")
            my_log.debug("-------------------------入网时间超过" + str(net_timeout) + "s,现在restart-------------------------------")
            t = -1
            return t
    else:
        t2 = time.time()
        t = format((t2 - t1), '.2f')
        if 'ctze' in config:
            ctze_time = getdata_line()
            t_ctze = time.time()
            while 'CTZE' not in ctze_time:
                ctze_time = getdata_line()
                if (time.time() - t_ctze) > 10:
                    break
            else:
                print(ctze_time)

        return t


def get_cell():
    global list
    """判断是否入网，发送AT+RADIO=CELL查看数据，返回cell信息"""
    global CELLID_error, AT_die, expected_cellid, AT_death, rsrp_old, rsrp_change_times
    # 打印串口缓冲的数据
    rsrp_change = 0
    getdata_0()
    send_AT0(str_CELL)
    data_cell = getdata_line()
    my_log.debug("读取" + str_CELL)
    try:
        i = 0
        while 'LUESTATS: CELL' not in data_cell:
            # send_AT0(str_CELL)
            if i < 5:
                data_cell = getdata_line()
            if 5>= i > 6:
                send_AT0(str_CELL)
                time.sleep(0.5)
                data_cell = getdata_0()
            if i >= 6:
                AT_death = AT_death + 1
                data_cell = "b'\r\nLUESTATS: CELL,0,0,0,0,0,-0,25\r\n\r\nOK\r\n'"
            i += 1

        else:
            rsrp_s = data_cell.rstrip("\\r\\n\\r\\nOK\\r\\n'").split(',')
            my_log.debug("发送AT+LUESTATS=CELL查询分割后输出：" + data_cell)
            rsrp = rsrp_s
        # 计算SNR
        # print(rsrp)
        if len(rsrp) == 8:
            SNR_match = re.match(r'-?\d+', rsrp[7])
            # SNR_int = int(SNR_match.group(0)) / 10
            SNR_int = int(SNR_match.group(0))
            SNR = SNR_int
            if SNR == '255':
                SNR = '26'
            my_log.debug(rsrp)
            # 判断选择小区是否跟上次一样
            pre_cellid = rsrp[2]
            rsrp_new = int(rsrp[4])
            if rsrp_old != 0:
                rsrp_change = rsrp_new - rsrp_old
            rsrp_old = rsrp_new

            if expected_cellid != pre_cellid:
                if expected_cellid == "504":
                    expected_cellid = pre_cellid
                else:
                    CELLID_error = CELLID_error + 1
                    my_log.debug("选择到不同小区")
                    expected_cellid = pre_cellid
                    # ["b'LUESTATS: CELL", '2505', '327', '1', '-84', '-88', '-73', '5']
                    if rsrp_change < -6:
                        rsrp_change_times += 1
        else:
            SNR = 25

        list.append(rsrp[1])
        list.append(rsrp[2])
        list.append(rsrp[3])
        list.append(rsrp[4])
        list.append(str(SNR))
    except:
        my_log.debug("--------------------------------AT+LUESTATS=CELL查询异常-------------------------------")


def get_psm_cell():
    global list
    """判断是否入网，发送AT+RADIO=CELL查看数据，返回cell信息"""
    global CELLID_error, AT_die, expected_cellid, AT_death
    # 打印串口缓冲的数据
    try:
        getdata_0()
        if config["send_AT_cell"] == '0':
            list3.append('-1')
            list3.append('-1')
            list3.append('-1')
            list3.append('-1')
            list3.append('-1')
        else:
            send_AT0(str_CELL)
            data_cell = getdata_line()
            my_log.debug("读取" + str_CELL)

            i = 0
            while 'LUESTATS: CELL' not in data_cell:
                # send_AT0(str_CELL)
                if i < 5:
                    data_cell = getdata_line()
                if 5 >= i > 6:
                    send_AT0(str_CELL)
                    time.sleep(0.5)
                    data_cell = getdata_0()
                if i >= 6:
                    AT_death = AT_death + 1
                    data_cell = "b'\r\nLUESTATS: CELL,0,0,0,0,0,-0,25\r\n\r\nOK\r\n'"
                i += 1

            else:
                rsrp_s = data_cell.rstrip("\\r\\n\\r\\nOK\\r\\n'").split(',')
                my_log.debug("发送AT+LUESTATS=CELL查询分割后输出：" + data_cell)
                rsrp = rsrp_s
            # 计算SNR
            if len(rsrp) == 8:
                SNR_match = re.match(r'-?\d+', rsrp[7])
                # SNR_int = int(SNR_match.group(0)) / 10
                SNR_int = int(SNR_match.group(0))
                SNR = SNR_int
                if SNR == '255':
                    SNR = '26'
                my_log.debug(rsrp)
                # 判断选择小区是否跟上次一样
                pre_cellid = rsrp[2]
                if expected_cellid != pre_cellid:
                    if expected_cellid == "504":
                        expected_cellid = pre_cellid
                    else:
                        CELLID_error = CELLID_error + 1
                        my_log.debug("选择到不同小区")
                        expected_cellid = pre_cellid
            else:
                SNR = 25
            list3.append(rsrp[1])
            list3.append(rsrp[2])
            list3.append(rsrp[3])
            list3.append(rsrp[4])
            list3.append(str(SNR))
    except:
            my_log.debug("--------------------------------AT+LUESTATS=CELL查询异常-------------------------------")


def get_cell_csinfo():
    global list, at_CSINFO
    if at_CSINFO == '1':
        # 打印串口缓存数据
        getdata_0()
        send_AT0(str_CSINFO)
        cell_csinfo = getdata_line()
        # if i >= 3:
        # AT_death = AT_death + 1
        # data_cell = "b'\r\nLUESTATS: CELL,0,0,0,0,0,-0,9999\r\n\r\nOK\r\n'"
        i = 0
        while 'CSINFO' not in cell_csinfo:
            if i < 5:
                cell_csinfo = getdata_line()
            if 5 >= i > 6:
                print("发送" + str_CSINFO + "命令超时，等5s再发送1次")
                my_log.debug("发送" + str_CSINFO + "命令超时，等5s再发送1次")
                # time.sleep(5)
                # if i > 5:
                send_AT0(str_CSINFO)
                time.sleep(0.5)
                cell_csinfo = getdata_line()
            if i >= 6:
                AT_death = AT_death + 1
                cell_csinfo = "CSINFO0,0,0,0,0,0,0"
            i = i + 1
        else:
            CSINFO = cell_csinfo.rstrip("\\r\\n\\r\\nOK\\r\\n'").split(',')
            my_log.debug(CSINFO)

        if len(CSINFO) == 7:
            cs = CSINFO[6]
            data_match = re.match(r'-?\d+', cs)
            CSINFO_data = data_match.group()
            # print(CSINFO_data)
            my_log.debug('查询CSINFO信息：')
            my_log.debug(CSINFO)
        else:
            CSINFO = [0, 0, 0, 0, 0, 0, 0]
        list.append(CSINFO[2])
        list.append(CSINFO[3])
        list.append(CSINFO[4])
        list.append(CSINFO[5])
        list.append(CSINFO_data)
    else:
        CSINFO = [0, 0, 0, 0, 0, 0, 0]
        list.append(CSINFO[2])
        list.append(CSINFO[3])
        list.append(CSINFO[4])
        list.append(CSINFO[5])
        list.append(0)


def get_NBSTATS():
    global list,list1
    """判断是否入网，发送AT+NBSTATS查看数据"""
    global CELLID_error, AT_die, expected_cellid, AT_death
    # 打印串口缓冲的数据
    getdata_0()
    send_AT0(str_NBSTATS)
    # data_cell = getdata_line()
    my_log.debug("读取" + str_NBSTATS)
    # try:
    i = 0
    # while 'LUESTATS: CELL' not in data_cell:
    # print('打印NBSTATS信息：')
    while i < 26:
        data_cell = getdata_line()
        # print(data_cell)
        if 'AccessReqcount' in data_cell:
            data_split = data_cell.split(": ")
            data_match = re.match(r'-?\d+', data_split[1]).group()
            list1.append(data_match)
        elif 'RARFailcount' in data_cell:
            data_split = data_cell.split(": ")
            data_match = re.match(r'-?\d+', data_split[1]).group()
            list1.append(data_match)
        elif 'Msg4Failcount' in data_cell:
            data_split = data_cell.split(": ")
            data_match = re.match(r'-?\d+', data_split[1]).group()
            list1.append(data_match)
        elif 'ContentionFailcount' in data_cell:
            data_split = data_cell.split(": ")
            data_match = re.match(r'-?\d+', data_split[1]).group()
            list1.append(data_match)
        elif 'AuthREQcoun' in data_cell:
            data_split = data_cell.split(": ")
            data_match = re.match(r'-?\d+', data_split[1]).group()
            list1.append(data_match)
        elif 'AuthFailcount' in data_cell:
            data_split = data_cell.split(": ")
            data_match = re.match(r'-?\d+', data_split[1]).group()
            list1.append(data_match)
        elif 'AuthRejectcount' in data_cell:
            data_split = data_cell.split(": ")
            data_match = re.match(r'-?\d+', data_split[1]).group()
            list1.append(data_match)
        elif 'AttachREQCount' in data_cell:
            data_split = data_cell.split(": ")
            data_match = re.match(r'-?\d+', data_split[1]).group()
            list1.append(data_match)
        elif 'AttachRejectcount' in data_cell:
            data_split = data_cell.split(": ")
            data_match = re.match(r'-?\d+', data_split[1]).group()
            list1.append(data_match)
        # 三个数据需要匹配
        #  rsrp: -77 -111 -73
        elif 'rsrp' in data_cell:
            data_split = data_cell.split(": ")
            list1 += re.findall(pattern=r"-?\d+", string=data_split[1])
        # 三个数据需要匹配
        elif 'snr' in data_cell:
            data_split = data_cell.split(": ")
            list1 += re.findall(pattern=r"-?\d+", string=data_split[1])
        # 三个数据需要匹配
        elif 'TotalAGC' in data_cell:
            data_split = data_cell.split(": ")
            list1 += re.findall(pattern=r"-?\d+", string=data_split[1])
        # 三个数据需要匹配
        elif 'PwrRACH' in data_cell:
            data_split = data_cell.split(": ")
            list1 += re.findall(pattern=r"-?\d+", string=data_split[1])
        # 三个数据需要匹配
        elif 'PwrPusch' in data_cell:
            data_split = data_cell.split(": ")
            list1 += re.findall(pattern=r"-?\d+", string=data_split[1])
        elif 'csFailCnt' in data_cell:
            data_split = data_cell.split(": ")
            data_match = re.match(r'-?\d+', data_split[1]).group()
            list1.append(data_match)
        elif 'mibFailCnt' in data_cell:
            data_split = data_cell.split(": ")
            data_match = re.match(r'-?\d+', data_split[1]).group()
            list1.append(data_match)
        elif 'sib1FailCnt' in data_cell:
            data_split = data_cell.split(": ")
            data_match = re.match(r'-?\d+', data_split[1]).group()
            list1.append(data_match)
        elif 'siFailCnt' in data_cell:
            data_split = data_cell.split(": ")
            data_match = re.match(r'-?\d+', data_split[1]).group()
            list1.append(data_match)
        elif 'npdschDecodeFailCnt' in data_cell:
            data_split = data_cell.split(": ")
            data_match = re.match(r'-?\d+', data_split[1]).group()
            list1.append(data_match)
        elif 'npdcchDecodeFailCnt' in data_cell:
            data_split = data_cell.split(": ")
            data_match = re.match(r'-?\d+', data_split[1]).group()
            list1.append(data_match)
        elif 'csUseRegInfoCnt' in data_cell:
            data_split = data_cell.split(": ")
            data_match = re.match(r'-?\d+', data_split[1]).group()
            list1.append(data_match)
        elif 'csNotUseRegInfoCnt' in data_cell:
            data_split = data_cell.split(": ")
            data_match = re.match(r'-?\d+', data_split[1]).group()
            list1.append(data_match)
        elif 'retransmitCnt' in data_cell:
            data_split = data_cell.split(": ")
            data_match = re.match(r'-?\d+', data_split[1]).group()
            list1.append(data_match)
        i = i + 1


def get_LWIPDATA():
    global list, list2
    """判断是否入网，发送AT+LWIPDATA查看数据"""
    global CELLID_error, AT_die, expected_cellid, AT_death
    # 打印串口缓冲的数据
    getdata_0()
    send_AT0(str_LWIPDATA)
    # data_cell = getdata_line()
    my_log.debug("读取" + str_LWIPDATA)
    # try:
    i = 0
    # while 'LUESTATS: CELL' not in data_cell:
    # print('打印LWIPDATA信息：')
    while i < 10:
        data_cell = getdata_line()
        # print(data_cell)
        if 'TxPacketCnt' in data_cell:
            data_split = data_cell.split(": ")
            data_match = re.match(r'-?\d+', data_split[1]).group()
            list2.append(data_match)
        elif 'RxPacketCnt' in data_cell:
            data_split = data_cell.split(": ")
            data_match = re.match(r'-?\d+', data_split[1]).group()
            list2.append(data_match)
        elif 'TxLossPacketCnt' in data_cell:
            data_split = data_cell.split(": ")
            data_match = re.match(r'-?\d+', data_split[1]).group()
            list2.append(data_match)
        elif 'TxMaxPacketSize' in data_cell:
            data_split = data_cell.split(": ")
            data_match = re.match(r'-?\d+', data_split[1]).group()
            list2.append(data_match)
        elif 'RxMaxPacketSize' in data_cell:
            data_split = data_cell.split(": ")
            data_match = re.match(r'-?\d+', data_split[1]).group()
            list2.append(data_match)
        elif 'TxPacketByte' in data_cell:
            data_split = data_cell.split(": ")
            data_match = re.match(r'-?\d+', data_split[1]).group()
            list2.append(data_match)
        elif 'RxPacketByte' in data_cell:
            data_split = data_cell.split(": ")
            data_match = re.match(r'-?\d+', data_split[1]).group()
            list2.append(data_match)
        i = i + 1

    # if len(CSINFO) == 7:
    #     cs = CSINFO[6]
    #     data_match = re.match(r'-?\d+', cs)
    #     CSINFO_data = data_match.group()
    #     # print(CSINFO_data)
    #     my_log.debug('查询CSINFO信息：')
    #     my_log.debug(CSINFO)


def get_cell_info():
    """判断是否入网，发送AT+RADIO=CELL查看数据，返回cell信息"""
    global CELLID_error, AT_die, expected_cellid, AT_death
    # 打印串口缓冲的数据
    getdata_0()
    send_AT0(str_CELL)
    time.sleep(0.5)
    data_cell = getdata_0()
    try:
        i = 0
        while data_cell in "b''":
            my_log.debug("读取" + str_CELL + "数据异常")
            print("读取" + str_CELL + "数据异常,休眠10s再次查询")
            time.sleep(1)
            send_AT0(str_CELL)
            time.sleep(0.5)
            data_cell = getdata_0()
            i += 1
            if i >= 2:
                AT_death = AT_death + 1
                data_cell = "b'\r\nLUESTATS: CELL,0,0,0,0,0,-0,99\r\n\r\nOK\r\n'"

        else:
            rsrp_s = data_cell.rstrip("\\r\\n\\r\\nOK\\r\\n'").split(',')
            my_log.debug("发送AT+LUESTATS=CELL查询分割后输出：" + data_cell)
            rsrp = rsrp_s
        # 计算SNR

        if len(rsrp) == 8:
            SNR_match = re.match(r'-?\d+', rsrp[7])
            # SNR_int = int(SNR_match.group(0)) / 10
            SNR_int = int(SNR_match.group(0))
            SNR = SNR_int
            my_log.debug(rsrp)
            # 判断选择小区是否跟上次一样
            pre_cellid = rsrp[2]
            if expected_cellid != pre_cellid:
                if expected_cellid == "504":
                    expected_cellid = pre_cellid
                else:
                    CELLID_error = CELLID_error + 1
                    my_log.debug("选择到不同小区")
                    expected_cellid = pre_cellid


        else:
            SNR = 99
    except:
        my_log.debug("--------------------------------AT+LUESTATS=CELL查询异常-------------------------------")

    # 发送AT+LUESTATS=CSINFO
    time.sleep(1)  # 两条查询命令间隔1s
    try:
        send_AT0(str_CSINFO)
        time.sleep(0.5)
        cell_csinfo = getdata_0()
        j = 0
        while 'CSINFO' not in cell_csinfo:
            my_log.debug("读取AT+LUESTATS=CSINFO数据异常")
            print("读取AT+LUESTATS=CSINFO数据异常,休眠10s再次查询")
            time.sleep(10)
            send_AT0(str_CSINFO)
            time.sleep(0.5)
            # cell_csinfo = None
            cell_csinfo = getdata_0()
            j += 1
            if j >= 2:
                AT_death = AT_death + 1
                cell_csinfo = "CSINFO0,0,0,0,0,0,0"
                # break
        else:
            time.sleep(1)
            # print(cell_csinfo)
            CSINFO = cell_csinfo.rstrip("\\r\\n\\r\\nOK\\r\\n'").split(',')
            my_log.debug(CSINFO)

        if len(CSINFO) == 7:
            cs = CSINFO[6]
            data_match = re.match(r'-?\d+', cs)
            CSINFO_data = data_match.group()
            # print(CSINFO_data)
            my_log.debug('查询CSINFO信息：')
            my_log.debug(CSINFO)
        else:
            CSINFO = [0, 0, 0, 0, 0, 0, 0]
    except:
        my_log.debug("--------------------------------AT+LUESTATS=CSINFO查询异常-------------------------------")
    try:
        list = []
        list.append(rsrp[1])
        list.append(rsrp[2])
        list.append(rsrp[3])
        list.append(rsrp[4])
        # list.append(rsrp[5])
        # list.append(rsrp[6])
        list.append(str(SNR))
        # list.append(t)
        list.append(CSINFO[2])
        list.append(CSINFO[3])
        list.append(CSINFO[4])
        list.append(CSINFO[5])
        list.append(CSINFO_data)
    except:
        my_log.debug("-------------------------------------写入数据异常-----------------------------------------")
    return list


def re_start():
    global t1, restart_times, restart_reboot, restart_power, task_end_time, psm_mode
    restart_times += 1

    if power == '1':
        send_AT0(str_REBOOT)
        restart_reboot += 1
        # 发送reboot时等待了0.5s，这里校准开始时间
        my_log.debug("-----------------------开始计算起始时间---------------------")
        t1 = time.time() - 0.5
        getdata_0()
        # print("restart前t1：" + str(t1))
    elif power == '4':
        my_log.debug("-----------------------开始计算起始时间---------------------")
        getdata_0()
        t1 = time.time()
        if psm_mode == '1':
            send_AT0("AT+REBOOT")
            restart_reboot += 1
        getdata_0()
    elif power == '2':
        enble_0_1()
        # restart_power += 1
        my_log.debug("-----------------------开始计算起始时间---------------------")
        t1 = time.time()
    elif power == '3':
        enble_0_1()
    if CTwing == 1:
        time.sleep(5)
        print("配置云平台参数" + str(CTwing_case_num))
        CTWing_case_config(CTwing_case_num, imei)
        if flash == 0:
            send_AT0(str_REBOOT)
            restart_reboot += 1
            time.sleep(2)

    if psm_mode == '1':
        time.sleep(1)
        print("reboot后等待入网........")
        t = get_net_time()
        print("reboot后的入网时间：" + str(t))
        my_log.debug("reboot后的入网时间：" + str(t))
        task_end_time = time.time()


def cfun_0_1():
    global AT_die, detach_time, t1, at_lsdata
    restart = 0
    if at_lsdata == "1":
        print("发送" + str_LSDATA)
        send_AT0(str_LSDATA)
    time.sleep(0.5)
    getdata_0()
    t2 = time.time()

    send_AT0("AT+CFUN=0")
    print('发送AT+CFUN=0')
    i = 0
    my_log.debug("等待CFUN0响应OK：")
    data = getdata_line()
    while 'OK' not in data:
        # my_log.debug("等待CFUN0响应：" + data)
        print(".", end='', flush=True)
        data = getdata_line()
        i += 1
        if i % 20 == 0:
            AT_die = AT_die + 1
            detach_time = -1
            re_start()
            # 标识restart
            restart = 1
            break
            # my_log.debug("等待CFUN0响应100s,现在restart，等待2s后再发CFUN0：" + data)
            # time.sleep(2)
            # send_AT0("AT+CFUN=0")
    else:
        t3 = time.time()
        detach_time = format((t3 - t2), '.2f')
        time.sleep(2)
    if restart == 1:
        t1 = t1
    else:
        send_AT0("AT+CFUN=1")
        t1 = time.time()
        my_log.debug("-----------------------开始计算起始时间---------------------")
        print('发送CFUN1等待入网')
        my_log.debug('发送CFUN1等待入网')


def enble_0_1():
    global AT_die, detach_time, t1, power, restart_power
    restart_power += 1
    getdata_0()
    if power == '3':
        if at_lsdata == "1":
            print("发送" + str_LSDATA)
            send_AT0(str_LSDATA)
        time.sleep(0.5)
        t2 = time.time()
        send_AT0("AT+CFUN=0")
        print('发送AT+CFUN=0')
        i = 0
        my_log.debug("等待CFUN0响应：")
        data = getdata_line()
        while 'OK' not in data:
            my_log.debug("等待CFUN0响应：" + data)
            print(".", end='', flush=True)
            data = getdata_line()
            i += 1
            if i % 20 == 0:
                AT_die = AT_die + 1
                detach_time = -1
                print("CFUN0等待超时50s")
                my_log.debug("CFUN0等待超时20s")
                break
        else:
            t3 = time.time()
            detach_time = format((t3 - t2), '.2f')
    # 休眠2s，防止鉴权失败问题
    time.sleep(2)
    ser_control.write((bytes(str_disable_ldo + "\r\n", encoding="utf-8")))
    print(str_disable_ldo)
    my_log.debug(str_disable_ldo)
    time.sleep(2)
    t1 = time.time()
    ser_control.write((bytes(str_enable_ldo + "\r\n", encoding="utf-8")))
    my_log.debug("-----------------------开始计算起始时间---------------------")
    print(str_enable_ldo)
    my_log.debug(str_enable_ldo)


def reboot():
    global AT_die, detach_time, t1, power, restart_reboot
    getdata_0()
    if power == '4':
        if at_lsdata == "1":
            print("发送" + str_LSDATA)
            send_AT0(str_LSDATA)
        time.sleep(0.5)
        t2 = time.time()
        if 'at_cfun' in config:
            print('--------------------------------------------------------')
            if j == 1:
                if config['learfcn0'] != '0':
                    send_AT0(config['learfcn0'])
            t3 = time.time()
            detach_time = format((t3 - t2), '.2f')
        else:
            send_AT0("AT+CFUN=0")
            print('发送AT+CFUN=0')
            i = 0
            my_log.debug("等待CFUN0响应：")
            data = getdata_line()
            while 'OK' not in data:
                my_log.debug("等待CFUN0响应：" + data)
                print(".", end='', flush=True)
                data = getdata_line()
                i += 1
                if i % 20 == 0:
                    AT_die = AT_die + 1
                    detach_time = -1
                    print("CFUN0等待超时20s")
                    my_log.debug("CFUN0等待超时20s")
                    break
            else:
                if j == 1:
                    if config['learfcn0'] != '0':
                        send_AT0(config['learfcn0'])
                t3 = time.time()
                detach_time = format((t3 - t2), '.2f')
    time.sleep(2)
    ser.write((bytes(str_REBOOT + "\r\n", encoding="utf-8")))
    print(str_REBOOT)
    restart_reboot += 1
    my_log.debug(str_REBOOT)
    t1 = time.time()
    my_log.debug("-----------------------开始计算起始时间---------------------")
    # getdata_0()
    # time.sleep(1)
    getdata_sleep(3)


def get_ota_data():
    send_AT("AT+CSDATA=1")
    time.sleep(2)
    # send_AT0("AT+CSDATA=2")
    # time.sleep(2)
    # send_AT0("AT+CSDATA=3")
    # time.sleep(2)
    send_AT("AT+CSDATA=0")


def do_ping(judge):
    global list, ping_1, ping_2, ping_all, ping_total_all, CEREG_0
    if judge == 1:
        # send_AT0("AT+LPING=172.22.1.201,3000,60,10") # 221.229.214.202
        getdata_0()
        if 'AT_CEREG' in config:
            send_AT0("AT+CEREG?")
            lines = getdata_line()
        else:
            lines = 'CEREG: 0,1'
        i = 0
        while 'CEREG: 0,1' not in lines and 'CEREG: 1,1' not in lines:
            lines = getdata_line()
            # print(lines)
            i += 1
            if i%5 == 1:
                send_AT0("AT+CEREG?")
            if i > 11:
                CEREG_0 += 1
                print("网络掉线：" + str(CEREG_0))
                my_log.debug("网络掉线：" + str(CEREG_0))
                # 对齐ping结果
                list.append('0')
                # 对齐丢包结果
                list.append('-1')
                # 对齐最小时延
                list.append('0')
                # 最大时延
                list.append('0')
                # 平均时延
                list.append('0')
                return '0'
    # else:
        ping_all = ping_all + 10
        t_ping = time.time()
        send_AT0(config['AT_LPING'])
        pattern = r"\d+.\d+.\d+,(\d+),.\d+,\d+"
        pattern1 = r"\d+.\d+,\d+,\d+,(\d+)"
        print(config['AT_LPING'])
        # AT+LPING=221.229.214.202,40000,200,10
        match = int(re.findall(pattern, config['AT_LPING'])[0])/1000
        match1 = re.findall(pattern1, config['AT_LPING'])[0]
        ping_total_all = int(match1)
        # print(match1)
        print("正在ping业务.....")
        data = getdata_line()
        # my_log.debug(data)
        reg_read_data = 'LPING: ' + match1 + '\S{1,}'
        reg_data = re.search(reg_read_data, data)
        i_time = time.time() - t_ping
        # print("等待超时时间" + str(i_time))
        # print(match*(int(match1)+2))
        while reg_data == None:
            if i_time < match*(int(match1)+1):
                if "LPING: 2\\r\\n" in data:
                    ping_2 = ping_2 + 1
                if "LPING: 1\\r\\n" in data:
                    ping_1 = ping_1 + 1
                data = getdata_line()
                reg_data = re.search(reg_read_data, data)
                i_time = time.time() - t_ping
                # i = i + 1
            else:
                ping_total = '0'
                ping_loss = '-1'
                list.append(ping_total)
                list.append(ping_loss)
                # 补齐
                list.append('5000')
                list.append('0')
                list.append('500')
                # print("等待超时时间" + str(i_time))
                print("ping匹配数据" + str(match*(int(match1)+1)) + "s异常现在断电重启")
                # re_start()
                return ping_total
        ping_total = reg_data.group().rstrip("\\r\\n'")
        print(ping_total)
        my_log.debug('+LPING: 1: ' + str(ping_1))
        my_log.debug('+LPING: 2: ' + str(ping_2))
        list.append(ping_total)
        ping_loss_d = ping_total.split(',')
        ping_loss = ping_loss_d[2]
        if ping_loss_d[4] == '0':
            list.append(ping_loss)
            list.append('9999')
            list.append('0')
            list.append('500')
        else:
            # 丢包数据
            list.append(ping_loss)
            # 最小时延
            list.append(ping_loss_d[4])
            # 最大时延
            list.append(ping_loss_d[5])
            # 平均时延
            list.append(ping_loss_d[6])
        return ping_total
    else:
        # 对齐ping结果
        list.append('0')
        # 对齐丢包结果
        list.append('-1')
        # 对齐最小时延
        list.append('0')
        # 最大时延
        list.append('0')
        # 平均时延
        list.append('0')
        return '0'


def write_excel(list):
    # 新建工作薄
    work_book = xlrd.open_workbook(r'C:\Users\letswin\Desktop\python_script\Test.xls')

    # 获取sheet
    # work_sheet = work_book.sheet_by_name('Test')
    work_wb = copy(work_book)
    work_sheet = work_wb.get_sheet('Test')
    for i in range(len(list)):
        work_sheet.write(2, i + 1, list[i])


def set_IP(addr):
    ser.flushInput()
    time.sleep(10)
    send_AT0('AT+LCGDDFCONT=' + addr)
    time.sleep(2)
    addr_data = str(ser.read_all())
    my_log.debug(addr_data)
    if 'OK' in addr_data:
        re_start()
        return addr


# 1、UE进入深睡
# 2、输出进入深睡等待时间
# 3、ping 5次，判断5次是否成功，全部成功返回“PASS”，有丢包返回“FAIL”
def DEEP_SLEEP():
    psm_state = psm_mode_1()
    while psm_state != 'error':
        my_log.debug(psm_state)
        k = 0
        print("等待进入深睡：", flush=True, end='')
        while "DEEP SLEEP" not in getdata():
            k = k + 1
            time.sleep(1)
            print('.', end='', flush=True)
            my_log.debug("等待休眠" + str(k) + "秒")
            if k > 150:
                print("等待深睡超过150秒")
                my_log.error("等待深睡超过150秒")
                my_log.error("Fail")
                # 解除PSM模式
                my_log.debug(psm_mode_0())
                return "Fail"
        else:
            print("\n" + "进入深睡，已经等待" + str(k) + "秒")
            my_log.error("进入深睡等待" + str(k) + "秒")
            time.sleep(1)
            # 输入唤醒命令：AT
            send_AT0(str_AT)
            # str_PING_2 = "AT+LPING=221.229.214.202,10000,60,2"
            # send_AT0(str_PING_2)
            # Initialized
            j = 0
            while "Initialized" not in getdata():
                j = j + 1
                time.sleep(1)
                if j >= 150:
                    print("ping超时" + str(j) + "秒")
                    my_log.error("ping超时" + str(j) + "秒")
                    my_log.error("Fail")
                    # 解除PSM模式
                    my_log.debug(psm_mode_0())
                    return "Fail"
            # while reg_data == None:
            #     data = getdata()
            #     reg_read_data = r'LPING: 5\S{1,}'
            #     reg_data = re.search(reg_read_data, data)
            #     time.sleep(1)
            #     j = j + 1
            #     my_log.debug("ping等待" + str(j) + "秒")
            #     if j >= 150:
            #         print("ping超时" + str(j) + "秒")
            #         my_log.error("ping超时" + str(j) + "秒")
            #         my_log.error("Fail")
            #         # 解除PSM模式
            #         my_log.debug(psm_mode_0())
            #         return "Fail"
            else:
                my_log.error("Pass")
                my_log.debug(psm_mode_0())
                return "Pass"
            # else:
            #     ping_total = reg_data.group().split(',')
            #     if len(ping_total) < 7:
            #         my_log.debug(ping_total)
            #         print(ping_total)
            #         print(psm_mode_0())
            #         return "Fail"
            #     fail_rate = ping_total[3]
            #     print("Ping失败率：" + fail_rate)
            #     my_log.error("Ping失败率：" + fail_rate)
            #     if float(fail_rate) == 100:
            #         return "Fail"
            #     elif float(fail_rate) <= 80.00:
            #         my_log.error("Pass")
            #         my_log.debug(psm_mode_0())
            #         return "Pass"
            #     else:
            #         my_log.error("Fail")
            #         my_log.debug(psm_mode_0())
            #         return "Fail"
    else:
        my_log.error("PSM状态：" + psm_state)
        print("PSM状态：" + psm_state)
        return "Fail"


# 登录电信云平台，成功返回('登录成功', '订阅成功')，否则失败
def CT_Wing():
    send_AT(str_LSERV)
    time.sleep(1)
    send_AT(str_LCTM2MINT)
    time.sleep(0.5)
    send_AT(str_REBOOT)
    time.sleep(2)
    i = 0
    # while "ADDR" not in getdata():
    while get_ip_addr() == 0:
        i = i + 2
        time.sleep(2)
        my_log.debug("正在等待入网...." + str(i) + "秒")
        if i > 100:
            print("入网时间超过100秒")
            my_log.error("入网时间超过100秒")
            my_log.error("Fail")
            return "Fail"
    else:
        print("入网成功")
        my_log.debug("入网成功")
    j = 0
    # ser.write(bytes(str_LCTM2MREG + "\r\n", encoding="utf-8"))
    # 等待自注册完成
    time.sleep(4)
    send_AT0(str_LCTM2MREG)
    # my_log.debug("发起注册，输入命令：" + str_LCTM2MREG)
    get_event = str(ser.readline())
    while "LWM2MEVENT: 0" not in get_event:
        j = j + 1
        time.sleep(1)
        get_event = str(ser.readline())
        my_log.debug("行读取" + get_event)
        if j > 30:
            print("等待登录超过30秒")
            my_log.error("等待登录超过30秒")
            my_log.error("Fail")
            return "Fail"
        # print("等待登录" + str(i) + "秒")
    else:
        Login = "登录成功"
        my_log.error("登录成功")
        print("登录成功")
    while "LWM2MEVENT: 2" not in get_event:
        time.sleep(1)
        j = j + 1
        get_event = str(ser.readline())
        my_log.debug("行读取" + get_event)
        my_log.debug("等待订阅" + str(j) + "秒")
        if j > 100:
            print("等待订阅时间超过100秒")
            my_log.error("等待订阅时间超过100秒")
            my_log.error('Fail')
            return "Fail"
    else:
        # print("订阅成功")
        my_log.error("订阅成功")
        Booking = "订阅成功"
        return "Pass"


def CTWing_Config(LCTM2MINIT, LSERV, LPSK):
    # send_AT('AT+LCTM2MINIT=869951046023283')
    send_AT(LCTM2MINIT)
    time.sleep(1)
    send_AT(LSERV)
    # AT+LSERV=221.229.214.202,5683'
    time.sleep(2)
    send_AT(LPSK)
    time.sleep(2)


def CTWing_case_config(num, imei):
    if num == 1:
        CT_UT_set()
    elif num == 2:
        CT_UT_DTLS_set()
    elif num == 3:
        CT_NonUT_JSON_set()
    elif num == 4:
        CT_NonUT_DTLS_JSON_set()
    elif num == 5:
        IP_ADDR = set_IP(addr="IP")
        if IP_ADDR == 'IP':
            CT_UT_set()
        else:
            print("设置IP栈失败")
            # return "设置IP栈失败", "设置IP栈失败", "设置IP栈失败"
    elif num == 6:
        IP_ADDR = set_IP(addr="IPV6")
        if IP_ADDR == 'IPV6':
            CT_UT_IPv6_set()
        else:
            print("设置IP栈失败")
            # return "设置IP栈失败", "设置IP栈失败", "设置IP栈失败"
    elif num == 7:
        IP_ADDR = set_IP(addr="IPV4V6")
        if IP_ADDR == 'IPV4V6':
            CT_UT_DTLS()
        else:
            print("设置IP栈失败")
            # return "设置IP栈失败", "设置IP栈失败", "设置IP栈失败"
    elif num == 8:
        CT_UT_DTLS_set()
    elif num == 9:
        CT_NonUT_JSON_set(imei)
    elif num == 10:
        CT_NonUT_DTLS_JSON_set()


# 电信——透传——明文——二进制,注册
def CT_UT_set():
    CTWing_Config(LCTM2MINIT='AT+LCTM2MINIT=869951046023283', LSERV='AT+LSERV=221.229.214.202,5683', LPSK=str_AT)


def CT_UT_IPv6_set():
    CTWing_Config(LCTM2MINIT='AT+LCTM2MINIT=869951046023283',
                  LSERV='AT+LSERV=240E:980:8120:28:84F4:C0C2:4A95:85F9,5683', LPSK=str_AT)


def CT_UT_DNS_set():
    CTWing_Config(LCTM2MINIT='AT+LCTM2MINIT=869951046023283', LSERV='AT+LSERV=lwm2m.ctwing.cn,5683', LPSK=str_AT)


# 电信——透传——DTLS——二进制
def CT_UT_DTLS_set():
    CTWing_Config(LCTM2MINIT='AT+LCTM2MINIT=862295048004637,0,2000,1', LSERV='AT+LSERV=221.229.214.202,5684',
                  LPSK='AT+LPSK=112233')


# 电信——非透传——明文——JSON
def CT_NonUT_JSON_set(imei):
    if imei == "Fail":
        # CTWing_Config(LCTM2MINIT='AT+LCTM2MINIT=862295048004611,1,2000,0', LSERV='AT+LSERV=221.229.214.202,5683', LPSK=str_AT) # 测试
        CTWing_Config(LCTM2MINIT='AT+LCTM2MINIT=862295048004611,1,2000,0', LSERV=config['LSERV'], LPSK=str_AT)
    else:
        CTWing_Config(LCTM2MINIT='AT+LCTM2MINIT=' + imei + ',' + config['LCTM2MINIT'], LSERV=config['LSERV'], LPSK=config['LPSK'])
    # CTWing_Config(LCTM2MINIT='AT+LCTM2MINIT=862295048015500,1,2000,0', LSERV='AT+LSERV=221.229.214.202,5683', LPSK=str_AT)   # 测试
    # CTWing_Config(LCTM2MINIT='AT+LCTM2MINIT=869951046023283(异常),1,2000,0', LSERV='AT+LSERV=221.229.214.202,5683', LPSK=str_AT) # 测试
    # CTWing_Config(LCTM2MINIT='AT+LCTM2MINIT=862295048000890,1,2000,0', LSERV='AT+LSERV=221.229.214.202,5683', LPSK=str_AT) # 杨万康
    # CTWing_Config(LCTM2MINIT='AT+LCTM2MINIT=862295048000783,1,2000,0', LSERV='AT+LSERV=221.229.214.202,5683', LPSK=str_AT) # 测试
    # CTWing_Config(LCTM2MINIT='AT+LCTM2MINIT=862295048025079,1,2000,0', LSERV='AT+LSERV=221.229.214.202,5683', LPSK=str_AT) # 测试
    # CTWing_Config(LCTM2MINIT='AT+LCTM2MINIT=862295048025087,1,2000,0', LSERV='AT+LSERV=221.229.214.202,5683', LPSK=str_AT) # 彭文敏
    # CTWing_Config(LCTM2MINIT='AT+LCTM2MINIT=862295048005287,1,2000,0', LSERV='AT+LSERV=221.229.214.202,5683', LPSK=str_AT) #外场 林树明 转彭文敏
    # CTWing_Config(LCTM2MINIT='AT+LCTM2MINIT=862295048007838,1,2000,0', LSERV='AT+LSERV=221.229.214.202,5683', LPSK=str_AT) #外场 严总 转彭文敏
    # CTWing_Config(LCTM2MINIT='AT+LCTM2MINIT=862295048009529,1,2000,0', LSERV='AT+LSERV=221.229.214.202,5683', LPSK=str_AT) #外场 刘双凤
    # CTWing_Config(LCTM2MINIT='AT+LCTM2MINIT=862295048007457,1,2000,0', LSERV='AT+LSERV=221.229.214.202,5683', LPSK=str_AT) #外场 杨总
    # CTWing_Config(LCTM2MINIT='AT+LCTM2MINIT=862295048010931,1,2000,0', LSERV='AT+LSERV=221.229.214.202,5683', LPSK=str_AT) #外场 刘勇
    # CTWing_Config(LCTM2MINIT='AT+LCTM2MINIT=862295048009677,1,2000,0', LSERV='AT+LSERV=221.229.214.202,5683', LPSK=str_AT) #外场 刘勇 转给吴海军测试
    # CTWing_Config(LCTM2MINIT='AT+LCTM2MINIT=862295048006475,1,2000,0', LSERV='AT+LSERV=221.229.214.202,5683', LPSK=str_AT) #外场 张峥
    # CTWing_Config(LCTM2MINIT='AT+LCTM2MINIT=862295048009453,1,2000,0', LSERV='AT+LSERV=221.229.214.202,5683', LPSK=str_AT) #外场 张晓燕
    # CTWing_Config(LCTM2MINIT='AT+LCTM2MINIT=862295048009032,1,2000,0', LSERV='AT+LSERV=221.229.214.202,5683', LPSK=str_AT) #外场 吴海军测试
    # CTWing_Config(LCTM2MINIT='AT+LCTM2MINIT=862295048009073,1,2000,0', LSERV='AT+LSERV=221.229.214.202,5683', LPSK=str_AT) #外场 李伟强
    # CTWing_Config(LCTM2MINIT='AT+LCTM2MINIT=862295048009420,1,2000,0', LSERV='AT+LSERV=221.229.214.202,5683', LPSK=str_AT) #外场 程检辉测试


# 电信——非透传——DTLS——JSON
def CT_NonUT_DTLS_JSON_set():
    CTWing_Config(LCTM2MINIT='AT+LCTM2MINIT=862295048004629,1,2000,1', LSERV='AT+LSERV=221.229.214.202,5684',
                  LPSK='AT+LPSK=3344')


def send_data_RAM():
    t1 = time.time()
    send_message_result = send_message(data='1122')
    t2 = time.time()
    t3 = t2 - t1
    t = format(t3, '.1f')
    if send_message_result != "Fail":
        return t
    else:
        return "发送数据失败"


# 注册云平台过程
def CTWing_case_run(judge, num):
    if judge == 1:
        if num == 1:
            print('W001透传_明文')
            my_log.debug('-----------------------------W001----------------------------------')
            reg_result = REG()
            send_data_result = "未发数据"
            dereg_result = DEREG()
            return reg_result[0], send_data_result, dereg_result[1]
        elif num == 2:
            print('W002透传_DTLS')
            # CT_UT_DTLS_set()
            reg_result = REG()
            send_data_result = "未发数据"
            dereg_result = DEREG()
            return reg_result[0], send_data_result, dereg_result
        elif num == 3:
            print('W003非透传_明文_JSON')
            my_log.debug('-----------------------------W003----------------------------------')
            # CT_NonUT_JSON_set()
            reg_result = REG()
            send_data_result = "未发数据"
            dereg_result = DEREG()
            return reg_result[0], send_data_result, dereg_result
        elif num == 4:
            print('W004非透传_DTLS_JSON')
            my_log.debug('-----------------------------W004----------------------------------')
            # CT_NonUT_DTLS_JSON_set()
            reg_result = REG()
            send_data_result = "未发数据"
            dereg_result = DEREG()
            return reg_result[0], send_data_result, dereg_result
        elif num == 5:
            print('W005透传_明文_IPV4_发送数据')
            my_log.debug('-----------------------------W005----------------------------------')
            IP_ADDR = set_IP(addr="IP")
            if IP_ADDR == 'IP':
                # CT_UT_set()
                reg_result = REG()
                if reg_result[1] == "订阅成功":
                    send_data_result = send_message(data='"1122"')
                    dereg_result = DEREG()
                    # time.sleep(2)
                    # send_result_b = send_message(data='1122')
                    return reg_result[0], send_data_result, dereg_result
                else:
                    set_IP(addr="IP")
                    dereg_result = DEREG()
                    return "订阅失败", "订阅失败", dereg_result
            else:
                return "设置IP栈失败", "设置IP栈失败", "设置IP栈失败"
        elif num == 6:
            print('W006透传_明文_IPV6_发送数据')
            my_log.debug('-----------------------------W006----------------------------------')
            IP_ADDR = set_IP(addr="IPV6")
            if IP_ADDR == 'IPV6':
                # CT_UT_IPv6_set()
                reg_result = REG()
                if reg_result[1] == "订阅成功":
                    send_data_result = send_message(data='"1122"')
                    dereg_result = DEREG()
                    time.sleep(1)
                    # send_result_b = send_message(data='1122')
                    return reg_result[0], send_data_result, dereg_result
                else:
                    set_IP(addr="IP")
                    return "订阅失败", "订阅失败", "订阅失败"
            else:
                return "设置IP栈失败", "设置IP栈失败", "设置IP栈失败"
        elif num == 7:
            print('W007透传_明文_域名_IPV4V6_发送数据')
            my_log.debug('-----------------------------W007----------------------------------')
            IP_ADDR = set_IP(addr="IPV4V6")
            if IP_ADDR == 'IPV4V6':
                # CT_UT_DTLS_set()
                reg_result = REG()
                if reg_result[1] == "订阅成功":
                    send_data_result = send_message(data='"3344"')
                    time.sleep(1)
                    dereg_result = DEREG()
                    set_IP(addr="IPV4V6")
                    return reg_result[0], send_data_result, dereg_result
                else:
                    set_IP(addr="IPV4V6")
                    dereg_result = DEREG()
                    return "订阅失败", "订阅失败", dereg_result
            else:
                print("设置IP栈失败")
                return "设置IP栈失败", "设置IP栈失败", "设置IP栈失败"
        elif num == 8:
            print('W008透传_DTLS_发送数据')
            my_log.debug('-----------------------------W008----------------------------------')
            # CT_UT_DTLS_set()
            reg_result = REG()
            if reg_result[1] == "订阅成功":
                send_data_result = send_message(data='"4455"')
                time.sleep(1)
                dereg_result = DEREG()
                # send_result_b = send_message(data='4455')
                # result1 = send_result_b + send_result_str
                # print(result)
                return reg_result[0], send_data_result, dereg_result
            else:
                dereg_result = DEREG()
                return "订阅失败", "订阅失败", dereg_result
        elif num == 9:
            print('W009非透传_明文_JSON_发送数据')
            my_log.debug('-----------------------------W009----------------------------------')
            reg_result = REG()
            if reg_result[1] == "订阅成功":
                time.sleep(1)
                # send_data_result = send_message(data='"{"pci":-9,"serviceId":2}"')
                send_data_result = send_message(
                    data='"{"id":"11:22:33:44:55:66","rssi":"[{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93}]""gps":"113.58761666,37.86236085","time":"1681182061"}"')
                time.sleep(1)
                if send_data_result == "发送数据30s超时":
                    dereg_result = DEREG_1()
                else:
                    dereg_result = DEREG()
                return reg_result[0], send_data_result, dereg_result[1]
            else:
                list.append("-1")
                dereg_result = DEREG_1()
                return "订阅失败", "订阅失败", dereg_result[1]
        elif num == 10:
            print('W010非透传_DTLS_JSON_发送数据')
            my_log.debug('-----------------------------W010----------------------------------')
            # CT_NonUT_DTLS_JSON_set()
            reg_result = REG()
            if reg_result[1] == "订阅成功":
                send_data_result = send_message(
                    data='"{"id":"11:22:33:44:55:66","rssi":"[{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93}]""gps":"113.58761666,37.86236085","time":"1681182061"}"')
                time.sleep(1)
                dereg_result = DEREG()
                # send_result_b = send_message(data='4455')
                # result1 = send_result_b + send_result_str
                # print(result)
                return reg_result[0], send_data_result, dereg_result
            else:
                dereg_result = DEREG_1()
                return "订阅失败", "订阅失败", dereg_result
        elif num == 11:
            print('W0011非透传_明文_JSON_发送数据')
            my_log.debug('-----------------------------W009----------------------------------')
            # 是否注册直接发送数据
            if config['send_data'] == 'nan_reg_sub':
                if j == 0:
                    reg_result = REG()
                    if reg_result[1] == "订阅成功":
                        time.sleep(1)
                        # send_data_result = send_message(data='"{"pci":-9,"serviceId":2}"')
                        send_message(data='"{"id":"11:22:33:44:55:66","rssi":"[{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93}]""gps":"113.58761666,37.86236085","time":"1681182061"}"')
                        # 对齐注销
                        list.append("-1")
                        time.sleep(1)

                else:
                    # 对齐
                    list.append("-1")
                    list.append("-1")
                    send_message(data='"{"id":"11:22:33:44:55:66","rssi":"[{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93}]""gps":"113.58761666,37.86236085","time":"1681182061"}"')
                    list.append("-1")
                    time.sleep(1)
            else:
                reg_result = REG()
                if reg_result[1] == "订阅成功":
                    time.sleep(1)
                    # send_data_result = send_message(data='"{"pci":-9,"serviceId":2}"')
                    send_data_result = send_message(
                        data='"{"id":"11:22:33:44:55:66","rssi":"[{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93}]""gps":"113.58761666,37.86236085","time":"1681182061"}"')
                    time.sleep(1)
                    if send_data_result == "发送数据30s超时":
                        dereg_result = DEREG_1()
                    else:
                        dereg_result = DEREG()
                    return reg_result[0], send_data_result, dereg_result[1]
                else:
                    list.append("-1")
                    dereg_result = DEREG_1()
                    return "订阅失败", "订阅失败", dereg_result[1]

    else:
        list.append('0')
        list.append('0')
        list.append('0')
        list.append('0')
        return 0, 0, 0, 0


# 电信——透传——明文——二进制,注册
def CT_UT(sers):
    send_AT('AT+LCTM2MINIT=869951046023283,,,')
    time.sleep(1)
    send_AT('AT+LSERV=' + sers)
    time.sleep(0.5)
    send_AT(str_REBOOT)
    time.sleep(2)
    i = 0
    while get_ip_addr() == 0:
        i = i + 2
        time.sleep(2)
        my_log.debug("正在等待入网...." + str(i) + "秒")
        if i > 100:
            print("入网时间超过100秒")
            my_log.error("入网时间超过100秒")
            my_log.error("Fail")
            # send_AT(str_REBOOT)
            return "Fail"
    else:
        print("入网成功")
        my_log.debug("入网成功")
    j = 0
    time.sleep(4)
    send_AT0(str_LCTM2MREG)
    get_event = str(ser.readline())
    while "LWM2MEVENT: 0" not in get_event:
        j = j + 1
        time.sleep(1)
        get_event = str(ser.readline())
        my_log.debug("行读取" + get_event)
        if j > 30:
            print("等待登录超过30秒")
            my_log.error("等待登录超过30秒")
            my_log.error("Fail")
            # send_AT(str_REBOOT)
            my_log.debug("发送命令：" + str_REBOOT)
            return "Fail"
        # print("等待登录" + str(i) + "秒")
    else:
        Login = "登录成功"
        my_log.error("登录成功")
        print("登录成功")
    while "LWM2MEVENT: 2" not in get_event:
        time.sleep(1)
        j = j + 1
        get_event = str(ser.readline())
        my_log.debug("行读取" + get_event)
        my_log.debug("等待订阅" + str(j) + "秒")
        if j > 100:
            print("等待订阅时间超过100秒")
            my_log.error("等待订阅时间超过100秒")
            my_log.error('Fail')
            # send_AT(str_REBOOT)
            return "Fail"
    else:
        # print("订阅成功")
        # send_AT(str_REBOOT)
        my_log.error("订阅成功")
        return "订阅成功"


# 电信——透传——DTLS——二进制
def CT_UT_DTLS():
    send_AT('AT+LSERV=221.229.214.202,5684')
    time.sleep(1)
    send_AT('AT+LCTM2MINIT=862295048004637,0,2000,1')
    # send_AT('AT+LCTM2MINIT=862295048004611,1,2000,0')
    time.sleep(0.5)
    send_AT('AT+LPSK=2233')
    time.sleep(1)
    send_AT(str_REBOOT)
    time.sleep(2)
    i = 0
    while get_ip_addr() == 0:
        i = i + 2
        time.sleep(2)
        my_log.debug("正在等待入网...." + str(i) + "秒")
        if i > 100:
            print("入网时间超过100秒")
            my_log.error("入网时间超过100秒")
            my_log.error("Fail")
            # send_AT(str_REBOOT)
            my_log.debug("发送命令：" + str_REBOOT)
            return "Fail"
    else:
        print("入网成功")
        my_log.debug("入网成功")
    j = 0
    time.sleep(4)
    send_AT0(str_LCTM2MREG)
    # my_log.debug("发起注册，输入命令：" + str_LCTM2MREG)
    get_event = str(ser.readline())
    while "LWM2MEVENT: 0" not in get_event:
        j = j + 1
        time.sleep(1)
        get_event = str(ser.readline())
        my_log.debug("行读取" + get_event)
        if j > 30:
            print("等待登录超过30秒")
            my_log.error("等待登录超过30秒")
            my_log.error("Fail")
            # send_AT(str_REBOOT)
            return "Fail"
        # print("等待登录" + str(i) + "秒")
    else:
        Login = "登录成功"
        my_log.error("登录成功")
        print("登录成功")
    while "LWM2MEVENT: 2" not in get_event:
        time.sleep(1)
        j = j + 1
        get_event = str(ser.readline())
        my_log.debug("行读取" + get_event)
        my_log.debug("等待订阅" + str(j) + "秒")
        if j > 100:
            print("等待订阅时间超过100秒")
            my_log.error("等待订阅时间超过100秒")
            my_log.error('Fail')
            # send_AT(str_REBOOT)
            return "Fail"
    else:
        # print("订阅成功")
        my_log.error("订阅成功")
        # send_AT(str_REBOOT)
        return "订阅成功"


# 电信——非透传——明文——JSON
def CT_NonUT_JSON():
    send_AT('AT+LSERV=221.229.214.202,5683')
    time.sleep(1)
    send_AT('AT+LCTM2MINIT=862295048004611,1,2000,0')
    time.sleep(1)
    send_AT(str_REBOOT)
    time.sleep(2)
    i = 0
    while get_ip_addr() == 0:
        i = i + 2
        time.sleep(2)
        my_log.debug("正在等待入网...." + str(i) + "秒")
        if i > 100:
            print("入网时间超过100秒")
            my_log.error("入网时间超过100秒")
            my_log.error("Fail")
            # send_AT(str_REBOOT)
            return "Fail"
    else:
        print("入网成功")
        my_log.debug("入网成功")
    j = 0
    time.sleep(4)
    send_AT0(str_LCTM2MREG)
    get_event = str(ser.readline())
    while "LWM2MEVENT: 0" not in get_event:
        j = j + 1
        time.sleep(1)
        get_event = str(ser.readline())
        my_log.debug("行读取" + get_event)
        if j > 30:
            print("等待登录超过30秒")
            my_log.error("等待登录超过30秒")
            my_log.error("Fail")
            # send_AT(str_REBOOT)
            return "Fail"
        # print("等待登录" + str(i) + "秒")
    else:
        Login = "登录成功"
        my_log.error("登录成功")
        print("登录成功")
    while "LWM2MEVENT: 2" not in get_event:
        time.sleep(1)
        j = j + 1
        get_event = str(ser.readline())
        my_log.debug("行读取" + get_event)
        my_log.debug("等待订阅" + str(j) + "秒")
        if j > 100:
            print("等待订阅时间超过100秒")
            my_log.error("等待订阅时间超过100秒")
            my_log.error('Fail')
            # send_AT(str_REBOOT)
            return "Fail"
    else:
        # print("订阅成功")
        my_log.error("订阅成功")
        # send_AT(str_REBOOT)
        return "订阅成功"


# 电信——非透传——DTLS——JSON
def CT_NonUT_DTLS_JSON():
    send_AT('AT+LSERV=221.229.214.202,5684')
    time.sleep(1)
    send_AT('AT+LCTM2MINIT=862295048004629,1,2000,1')
    time.sleep(1)
    send_AT('AT+LPSK=3344')
    time.sleep(1)
    send_AT(str_REBOOT)
    time.sleep(2)
    i = 0
    while get_ip_addr() == 0:
        i = i + 2
        time.sleep(2)
        my_log.debug("正在等待入网...." + str(i) + "秒")
        if i > 100:
            print("入网时间超过100秒")
            my_log.error("入网时间超过100秒")
            my_log.error("Fail")
            # send_AT(str_REBOOT)
            return "Fail"
    else:
        print("入网成功")
        my_log.debug("入网成功")
    j = 0
    time.sleep(4)
    send_AT0(str_LCTM2MREG)
    get_event = str(ser.readline())
    while "LWM2MEVENT: 0" not in get_event:
        j = j + 1
        time.sleep(1)
        get_event = str(ser.readline())
        my_log.debug("行读取" + get_event)
        if j > 30:
            print("等待登录超过30秒")
            my_log.error("等待登录超过30秒")
            my_log.error("Fail")
            # send_AT(str_REBOOT)
            return "Fail"
        # print("等待登录" + str(i) + "秒")
    else:
        Login = "登录成功"
        my_log.error("登录成功")
        print("登录成功")
    while "LWM2MEVENT: 2" not in get_event:
        time.sleep(1)
        j = j + 1
        get_event = str(ser.readline())
        my_log.debug("行读取" + get_event)
        my_log.debug("等待订阅" + str(j) + "秒")
        if j > 100:
            print("等待订阅时间超过100秒")
            my_log.error("等待订阅时间超过100秒")
            my_log.error('Fail')
            # send_AT(str_REBOOT)
            return "Fail"
    else:
        # print("订阅成功")
        my_log.error("订阅成功")
        # send_AT(str_REBOOT)
        return "订阅成功"


def CT_Wing_MQTT():
    ser.write(bytes(str_LSERV_MQTT + "\r\n", encoding="utf-8"))
    time.sleep(0.2)
    ser.write(bytes(str_REBOOT + "\r\n", encoding="utf-8"))
    time.sleep(2)

    ser.write(bytes(str_LMQTTCON + "\r\n", encoding="utf-8"))
    time.sleep(0.2)
    i = 0
    while "ADDR" not in getdata():
        i = i + 2
        ser.write(bytes(str_ADDR + "\r\n", encoding="utf-8"))
        my_log.debug("输入命令： " + str_ADDR)
        # print("输入命令： " + str_ADDR)
        time.sleep(2)
        my_log.debug("正在等待入网...." + str(i) + "秒")
        # print("正在等待入网...." + str(i) + "秒")
        if i > 100:
            print("入网时间超过100秒")
            my_log.error("入网时间超过100秒")
            return "Fail"
    else:
        print("入网成功")
        my_log.debug("入网成功")
    while "OK" not in getdata():
        time.sleep(1)
        i = i + 1
        if i > 100:
            print("等待超时")
            break
    else:
        print("登录成功")
    str_SEND = 'AT + LMQTTPUB =“mqData”, "0", "{mqtt_test}:20"'
    ser.write(bytes(str_SEND + "\r\n", encoding="utf-8"))
    time.sleep(0.2)
    j = 0
    while "OK" not in getdata():
        j = j + 1
        time.sleep(1)
        if j > 100:
            print("发送超时")
        else:
            print("发送成功")
            return "Pass"


def wait_CTwing_reg():
    global CTwing_reg, CTwing_sub
    print("开始注册云平台.......")
    j = 0
    t1 = time.time()
    if auto_CTwing != "1":
        send_AT0(str_LCTM2MREG)
    # send_AT0(str_LCTM2MREG)
    get_event = getdata_line()
    while "LWM2MEVENT: 0" not in get_event:
        j = j + 1
        get_event = getdata_line()
        if j > 20:
            print("等待登录超过20秒")
            my_log.error("等待登录超过20秒")
            my_log.error("Fail")
            CTWing_case_config(CTwing_case_num, imei)
            time.sleep(2)
            CTwing_reg += 1
            list3.append("0")
            list3.append("-1")
            return "等待登录超时", "等待登录超时"
    else:
        my_log.error("登录成功")
        t3 = format((time.time() - t1), '.2f')
        list3.append(t3)
        print("登录成功")
    while "LWM2MEVENT: 2" not in get_event:
        j = j + 1
        get_event = getdata_line()
        if j > 20:
            print("等待订阅时间超过20秒")
            my_log.error("等待订阅时间超过20秒")
            CTWing_case_config(CTwing_case_num, imei)
            time.sleep(2)
            CTwing_sub += 1
            list3.append("-1")
            # my_log.debug("--------------------------debug事件2失败后写入excel------------------------")
            return "订阅超时", "订阅超时"
    else:
        my_log.error("订阅成功")
        print("订阅成功")
        t4 = format((time.time() - t1), '.2f')
        list3.append(t4)
        return t4, '订阅成功'


# 注册登录云平台
def REG():
    global CTwing_reg, CTwing_sub, auto_CTwing, psm_mode, imei
    print("开始注册云平台.......")
    j = 0
    t1 = time.time()
    if auto_CTwing != "1":
        send_AT0(str_LCTM2MREG)
    get_event = getdata_line()
    while "LWM2MEVENT: 0" not in get_event:
        j = j + 1
        get_event = getdata_line()
        if j > 30:
            print("等待登录超过30秒")
            my_log.error("等待登录超过30秒")
            my_log.error("Fail")
            CTWing_case_config(CTwing_case_num, imei)
            time.sleep(2)
            CTwing_reg += 1
            list.append("-1")
            list.append("-1")
            return "等待登录超时", "等待登录超时"
    else:
        my_log.error("登录成功")
        t3 = format((time.time() - t1), '.2f')
        list.append(t3)
        print("登录成功")
    while "LWM2MEVENT: 2" not in get_event:
        j = j + 1
        get_event = getdata_line()
        if j > 30:
            print("等待订阅时间超过30秒")
            my_log.error("等待订阅时间超过30秒")
            CTWing_case_config(CTwing_case_num, imei)
            time.sleep(2)
            CTwing_sub += 1
            list.append("-1")
            return "订阅超时", "订阅超时"
    else:
        my_log.error("订阅成功")
        print("订阅成功")
        t4 = format((time.time() - t1), '.2f')
        list.append(t4)
        return t4, '订阅成功'


def PSM_DEREG():
    getdata_0()
    t2 = time.time()
    send_AT0('AT+LCTM2MDEREG')
    dereg_data = getdata_line()
    i = 0
    while "+LWM2MEVENT: 3" not in dereg_data:
        dereg_data = getdata_line()
        i = i + 1
        while i > 30:
            t = -1
            list3.append("注销失败")
            print("注销失败")
            return "注销失败", "注销失败"
    else:
        t0 = format((time.time() - t2), '.1f')
        list3.append(t0)
        print("注销成功")
        return "注销成功", t0


# 注销登录云平台回复事件3
def DEREG():
    global psm_mode
    getdata_0()
    t2 = time.time()
    send_AT0('AT+LCTM2MDEREG')
    dereg_data = getdata_line()
    i = 0
    while "+LWM2MEVENT: 7" not in dereg_data:
        dereg_data = getdata_line()
        i = i + 1
        while i > 30:
            t = -1
            if psm_mode == '1':
                list3.append("-1")
            else:
                list.append("-1")
            print("注销失败")
            return "注销失败", "注销失败"
    else:
        t0 = format((time.time() - t2), '.1f')
        if psm_mode == "1":
            list3.append(t0)
        else:
            list.append(t0)
        print("注销成功")
        return "注销成功", t0


# 注销云平台回复“ok”
def DEREG_1():
    global psm_mode
    getdata_0()
    t2 = time.time()
    send_AT0('AT+LCTM2MDEREG')
    # my_log.debug("-----------DEREG_1-------------注销失败2")
    dereg_data = getdata_line()
    i = 0
    while "OK" not in dereg_data:
        dereg_data = getdata_line()
        i = i + 1
        # my_log.debug("-----------DEREG_1-------------注销失败3")
        if i > 15:
            if psm_mode == "1":
                list3.append("-1")
            else:
                list.append("-1")
            # re_start()
            t = -1
            print("注销失败")
            return "注销失败", "注销失败"
    else:
        t0 = format((time.time() - t2), '.1f')
        list.append(t0)
        if psm_mode == "1":
            list3.append(t0)
        print("注销成功")
        return "注销成功", t0


def W001():
    ''' 1、配置参数
        AT+LCTM2MINIT=869951046023283,,,
        AT+LSERV=221.229.214.202,5683
        2、发起注册（收到事件2判断为pass） '''
    print('W001透传_明文')
    my_log.debug('-----------------------------W001----------------------------------')
    result = CT_UT(sers='221.229.214.202,5683')
    my_log.debug(DEREG())
    # send_AT(str_REBOOT)
    return "订阅成功"


# 862295048004637 透传——DTLS
def W002_DTLS():
    ''' 1、配置参数
        AT+LSERV=221.229.214.202,5684
        AT+LCTM2MINIT=862295048004637,0,2000,1
        AT+LPSK=2233
        2、发起注册（收到事件2判断为pass） '''
    print('W002透传_DTLS')
    my_log.debug('-----------------------------W002----------------------------------')
    result = CT_UT_DTLS()
    my_log.debug(DEREG())
    return "订阅成功"


# 862295048004611 非透传——明文——JSON
def W003_JSON():
    ''' 1、配置参数
        AT+LSERV=221.229.214.202,5683
        AT+LCTM2MINIT=862295048004611,1,2000,0
        2、发起注册（收到事件2判断为pass） '''
    print('W003非透传_明文_JSON')
    my_log.debug('-----------------------------W003----------------------------------')
    result = CT_NonUT_JSON()
    my_log.debug(DEREG())
    return "订阅成功"


# 862295048004629 非透传——DTLS——JSON
def W004_DTLS_JSON():
    # ''' 1、配置参数：
    #     AT+LPSK=3344
    #     AT+LSERV=221.229.214.202,5684
    #     AT+LCTM2MINIT=862295048004629,1,2000,1
    #     2、发起注册（收到事件2判断为pass） '''
    print('W004非透传_DTLS_JSON')
    my_log.debug('-----------------------------W004----------------------------------')
    result = CT_NonUT_DTLS_JSON()
    my_log.debug(DEREG())
    return "订阅成功"


# 869951046023283 透传——明文 IPV4
def W005_IPV4_senddata():
    # ''' 1、配置参数：设置为IPV4
    #     AT+LCGDDFCONT=IP
    #     AT+LCTM2MINIT=869951046023283,,,
    #     AT+LSERV=lwm2m.ctwing.cn,5683
    #     2、发起注册
    #     3、发送字符串数据和十六进制数据
    #     AT+LCTM2MSEND="1122"
    #     AT+LCTM2MSEND=1122
    #     收到两个 +LWM2MEVENT: 5,4事件判断为”pass“ '''
    print('W005透传_明文_IPV4_发送数据')
    my_log.debug('-----------------------------W005----------------------------------')
    IP_ADDR = set_IP(addr="IP")
    if IP_ADDR == 'IP':
        result = CT_UT(sers='lwm2m.ctwing.cn,5683')
        if result == "订阅成功":
            send_result_str = send_message(data='"1122"')
            time.sleep(2)
            send_result_b = send_message(data='1122')
            result1 = send_result_b + send_result_str
            return result + result1
        else:
            set_IP(addr="IP")
            return "订阅失败"
    else:
        return "设置IP栈失败"


# 869951046023283 透传——明文 IPV6
def W006_IPV6_senddata():
    ''' 1、配置参数：设置为IPV6
        AT+LCGDDFCONT=IPV6
        AT+LCTM2MINIT=869951046023283,,,
        AT+LSERV=240E:980:8120:28:84F4:C0C2:4A95:85F9,5683
        2、发起注册
        3、发送字符串数据和十六进制数据
        AT+LCTM2MSEND="2233"
        AT+LCTM2MSEND=2233
        收到两个 +LWM2MEVENT: 5,4事件判断为”pass“ '''
    print('W006透传_明文_IPV6_发送数据')
    my_log.debug('-----------------------------W006----------------------------------')
    IP_ADDR = set_IP(addr="IPV6")
    if IP_ADDR == 'IPV6':
        result = CT_UT(sers='240E:980:8120:28:84F4:C0C2:4A95:85F9,5683')
        if result == "订阅成功":
            send_result_str = send_message(data='"2233"')
            time.sleep(2)
            send_result_b = send_message(data='2233')
            result1 = send_result_b + send_result_str
            my_log.debug(DEREG())
            set_IP(addr="IP")
            return result + result1
        else:
            set_IP(addr="IP")
            return "订阅失败"
    else:
        return "设置IP栈失败"


# 869951046023283 透传——明文 IPV4V6
def W007_IPV4V6_senddata():
    ''' 1、配置参数：设置为IPV4V6
        AT+LCGDDFCONT=IPV4V6
        AT+LCTM2MINIT=869951046023283,,,
        AT+LSERV=lwm2m.ctwing.cn,5683
        2、发起注册
        3、发送数据（字符串数据和十六进制）
        AT+LCTM2MSEND="3344"
        AT+LCTM2MSEND=3344
        收到两个 +LWM2MEVENT: 5,4事件判断为”pass“ '''
    print('W007透传_明文_IPV4V6_发送数据')
    my_log.debug('-----------------------------W007----------------------------------')
    IP_ADDR = set_IP(addr="IPV4V6")
    if IP_ADDR == 'IPV4V6':
        result = CT_UT(sers='lwm2m.ctwing.cn,5683')
        if result == "订阅成功":
            send_result_str = send_message(data='"3344"')
            time.sleep(2)
            send_result_b = send_message(data='3344')
            result1 = send_result_b + send_result_str
            # print(result)
            my_log.debug(DEREG())
            set_IP(addr="IP")
            return result1
        else:
            set_IP(addr="IP")
            return "订阅失败"
    else:
        print("设置IP栈失败")
        return "设置IP栈失败"


def W008_DTLS_senddata():
    ''' 1、配置参数
        AT+LSERV=221.229.214.202,5684
        AT+LCTM2MINIT=862295048004637,0,2000,1
        AT+LPSK=2233
        2、发起注册
        3、发送数据（串数据和十六进制数据）
        AT+LCTM2MSEND="4455"
        AT+LCTM2MSEND=4455 '''
    print('W008透传_DTLS_发送数据')
    my_log.debug('-----------------------------W008----------------------------------')
    result_1 = CT_UT_DTLS()
    if result_1 == "订阅成功":
        send_result_str = send_message(data='"4455"')
        time.sleep(2)
        send_result_b = send_message(data='4455')
        result1 = send_result_b + send_result_str
        # print(result)
        my_log.debug(DEREG())
        return result_1 + result1
    else:
        return "订阅失败"


def W009_JSON_senddata():
    ''' 1、配置参数
        AT+LSERV=221.229.214.202,5683
        AT+LCTM2MINIT=862295048004611,1,2000,0
        2、发起注册
        3、发送数据（串数据和十六进制数据）
        AT+LCTM2MSEND="{"pci":-9,"serviceId":2}" '''
    print('W009非透传_明文_JSON_发送数据')
    my_log.debug('-----------------------------W009----------------------------------')
    result_1 = CT_NonUT_JSON()
    if result_1 == "订阅成功":
        if 'send_data_byte' in config:
            if config['send_data_byte'] == '1':
            # result1 = send_message(data='"{"pci":-9,"serviceId":2}"')
                result1 = send_message(
                    data='"{"id":"11:22:33:44:55:66","rssi":"[{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93}]""gps":"113.58761666,37.86236085","time":"1681182061"}"')
            # print(result)
            elif config['send_data_byte'] == '2':
                result1 = send_message(data='{"id":"11:22","rssi":"[{major:00001,minor:00001,R:-93}]""gps":"113.61666,37.86","time":"1681182061"}')
            else:
                result1 = send_message(data='{"id":"1","r":"[{m:0}]"}')

        else:
            result1 = send_message(
                data='"{"id":"11:22:33:44:55:66","rssi":"[{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93},{major:00001,minor:00001,R:-93}]""gps":"113.58761666,37.86236085","time":"1681182061"}"')
        # print(result)
        my_log.debug(DEREG())
        return result_1 + result1
    else:
        return "订阅失败"


def W010_DTLS_JSON_senddata():
    ''' 1、配置参数：
        AT+LPSK=3344
        AT+LSERV=221.229.214.202,5684
        AT+LCTM2MINIT=862295048004629,1,2000,1
        2、发起注册
        3、发送数据（串数据和十六进制数据）
        "{"pci":-10,"serviceId":2}" '''
    print('W010非透传_DTLS_JSON_发送数据')
    my_log.debug('-----------------------------W010----------------------------------')
    result_1 = CT_NonUT_DTLS_JSON()
    if result_1 == "订阅成功":
        result1 = send_message(data='"{"pci":-10,"serviceId":2}"')
        # print(result)
        my_log.debug(DEREG())
        return result_1 + result1
    else:
        return "订阅失败"


def send_message(data):
    global CTwing_data, psm_mode
    send_AT("AT+LCTM2MSEND=1," + data)
    t1 = time.time()
    get_data = getdata_line()
    i = 0
    while 'LWM2MEVENT: 5: 4' not in get_data:
        i = i + 1
        get_data = getdata_line()
        if i > 30:
            my_log.debug("发送数据30s超时")
            if psm_mode == "1":
                list3.append("-1")
                my_log.debug("--------------------------debug发送数据失败后写入excel------------------------")
            else:
                list.append("-1")
            # DEREG_1()
            CTwing_data = CTwing_data + 1
            return "发送数据30s超时"
    else:
        t = format((time.time() - t1), '.1f')
        if psm_mode != '1':
            list.append(t)
        else:
            list3.append(t)
        my_log.debug("发送数据成功")
        print("发送数据成功")
        return t


def CTwing_test(judge, set):
    if judge == 1:
        print("开始云平台测试..........")
        my_log.debug("----------------------------开始云平台测试-----------------------------------")
        if set == 0:
            CTWing_result = W001()
            return CTWing_result
        if set == 1:
            CTWing_result = W002_DTLS()
            return CTWing_result
        if set == 2:
            CTWing_result = W003_JSON()
            return CTWing_result
        if set == 3:
            CTWing_result = W004_DTLS_JSON()
            return CTWing_result
        if set == 4:
            CTWing_result = W005_IPV4_senddata()
            return CTWing_result
        if set == 5:
            CTWing_result = W006_IPV6_senddata()
            return CTWing_result
        if set == 6:
            CTWing_result = W007_IPV4V6_senddata()
            return CTWing_result
        if set == 7:
            CTWing_result = W008_DTLS_senddata()
            return CTWing_result
        if set == 8:
            CTWing_result = W009_JSON_senddata()
            return CTWing_result
        if set == 9:
            CTWing_result = W010_DTLS_JSON_senddata()
            return CTWing_result

        # print(W001())
        # print(W002_DTLS())
        # print(W003_JSON())
        # print(W004_DTLS_JSON())
        # print(W005_IPV4_senddata())
        # print(W006_IPV6_senddata())
        # print(W007_IPV4V6_senddata())
        # print(W008_DTLS_senddata())
        # print(W009_JSON_senddata())
        # print(W010_DTLS_JSON_senddata())
    else:
        return "不测试云平台"


# plt.ion()
# 创建Figure和两个坐标.
figure, ax = plt.subplots(4, 1)
# return AxesImage object for using.
ax[0].set_autoscaley_on(True)
# ax.set_xlim(min_x, max_x)
ax[0].set_autoscaley_on(True)
ax00 = ax[0].twinx()
ax11 = ax[1].twinx()
ax22 = ax[2].twinx()

# ax.grid()
plt.rcParams["font.sans-serif"] = ["SimHei"]
plt.rcParams["axes.unicode_minus"] = False

def str_int(numbers):
    new_numbers = []
    for n in range(len(numbers)-2):
        data = float(numbers[n+2])
        new_numbers.append(data)
    return new_numbers

list_psm_active_y = []
list_psm_get_net_y = []


def draw_line_psm(excel_name):
    global j, drawline, psm_active, psm_get_net_time
    # work_book = xlrd.open_workbook(r'C:\Users\letswin\Desktop\python_script_cmw500_2CELL\IMEI\Test.xls')
    color = ['#e6194B', '#3cb44b', '#ffe119', '#4363d8',
             '#f58231', '#911eb4', '#42d4f4', '#f032e6',
             '#bfef45', '#fabed4', '#469990', '#dcbeff',
             '#9A6324', '#fffac8', '#800000', '#aaffc3',
             '#808000', '#ffd8b1', '#000075', '#a9a9a9',
             '#ffffff', '#000000']

    work_book = xlrd.open_workbook(path + '\\' + excel_name)
    # 读取sheet3异常统计数据，写入Y轴
    sheet_3 = work_book.sheet_by_index(3)
    data_sheet3_col1 = sheet_3.col_values(1)
    del data_sheet3_col1[0:2]
    # data_col2 = sheet_1.col_values(2)
    # del data_col2[0:2]
    # data_col3 = sheet_1.col_values(3)
    # del data_col3[0:2]
    # data_col4 = sheet_1.col_values(4)
    # del data_col4[0:2]

    # 读取sheet4 的数据写入Y轴
    sheet4 = work_book.sheet_by_index(4)
    data_sheet4_col1 = np.divide(str_int(sheet4.col_values(2)), 10)

    data_sheet4_col2 = str_int(sheet4.col_values(3))
    # RSRP
    data_sheet4_col7 = np.divide(str_int(sheet4.col_values(7)), 10)
    # SNR
    data_sheet4_col8 = str_int(sheet4.col_values(8))
    data_sheet4_col9 = str_int(sheet4.col_values(10))
    data_sheet4_col10 = str_int(sheet4.col_values(11))
    data_sheet4_col12 = str_int(sheet4.col_values(13))

    # col1_3_value = str_int(sheet.col_values(3))
    # del data_sheet0_col3[0:2]

    # X轴数据
    xdata = np.arange(len(data_sheet3_col1))
    # 使能交互模式
    plt.ion()
    # print(xdata)
    # print(list_psm_get_net_y)
    line1 = ax[0].plot(xdata, data_sheet3_col1, label="Abnormal_RST", color=color[0], linewidth=0.5)
    line2 = ax[0].plot(xdata, list_psm_get_net_y, label="PSM_net_timeout", color=color[4], linewidth=0.5)

    line3 = ax[0].plot(xdata, list_psm_active_y, label="AT_Sleep", color=color[5], linewidth=0.5)
    # line3 = ax[1].plot(xdata, data_col3, label="异常复位", color="blue", linewidth=1)  # , lable="异常复位"
    # line4 = ax[1].plot(xdata, data_col4, label="丢包次数", color="black", linewidth=1)  # , lable="丢包次数"

    # print(type(xdata))
    line5 = ax[1].plot(xdata, data_sheet4_col1, label="Detach_Sleep_time/10", color=color[1], linewidth=0.5)
    line6 = ax[1].plot(xdata, data_sheet4_col2, label="PSM_net_time", color=color[2], linewidth=0.5)
    line7 = ax[1].plot(xdata, data_sheet4_col10, label="Send_data_time", color=color[3], linewidth=0.5)
    line8 = ax[1].plot(xdata, data_sheet4_col9, label="CTwing_Reg_Sub", color=color[6], linewidth=0.5)

    line9 = ax[2].plot(xdata, data_sheet4_col7, label="RSRP/10", color=color[7], linewidth=0.5)
    line10 = ax[2].plot(xdata, data_sheet4_col8, label="SNR", color=color[8], linewidth=0.5)
    line11 = ax[3].plot(xdata, data_sheet4_col12, label="唤醒-业务-休眠时间", color=color[9], linewidth=0.5)
    # print(data_sheet4_col7)
    # print(data_sheet4_col8)

    title_name = str(time_excel) + '_' + str(test_port) + '_Test_Mode_' + str(power) + '_PSM'
    if j == 0:
        # plt.xlabel("次数")
        # plt.ylabel("异常")
        # plt.title(title_name, color='red', loc='center')
        plt.title(title_name, x=0.5, y=3.6, color='red')
        # ax[0,0].set_title(title_name, color='red', loc='center')
        # ax[0,0].set_ylabel("时间s", loc='center')
        ax[1].set_ylim(-5, 20)
        ax[1].set_yticks([-3, 0, 10, 15])
        # RSRP , SNR
        ax[2].set_ylim(-20, 30)
        ax[2].set_yticks([-8, 0, 10, 25])
        # print(plt)
        ax[0].legend(loc='upper right', fontsize=6)
        ax[1].legend(loc='upper right', fontsize=6)
        ax[2].legend(loc='upper right', fontsize=6)
        ax[3].legend(loc='upper right', fontsize=6)

    # Need both of these in order to rescale
    ax[1].relim()
    ax[1].autoscale_view()
    ax[0].relim()
    ax[0].autoscale_view()
    # draw and flush the figure .
    figure.canvas.draw()
    figure.canvas.flush_events()

    # plt.tight_layout()

    if drawline == '1':
        # print
        plt.show()
    # 关闭交互模式
    plt.ioff()
    plt.savefig(path + '\\' + title_name + '.png', dpi=300, bbox_inches='tight')


def draw_line(excel_name):
    global j, drawline
    # time.sleep(120)
    excel_time = time.time()
    color = ['#e6194B', '#3cb44b', '#ffe119', '#4363d8',
           '#f58231', '#911eb4', '#42d4f4', '#f032e6',
           '#bfef45', '#fabed4', '#469990', '#dcbeff',
           '#9A6324', '#fffac8', '#800000', '#aaffc3',
           '#808000', '#ffd8b1', '#000075', '#a9a9a9',
           '#ffffff', '#000000']
    # 打开excel
    work_book = xlrd.open_workbook(path + '\\' + excel_name)

    # 读取sheet3异常统计数据，写入Y轴
    sheet_1 = work_book.sheet_by_index(3)
    # 获取失败次数
    data_col1 = sheet_1.col_values(1)
    del data_col1[0:2]
    # 获取云平台登录失败次数
    data_col2 = sheet_1.col_values(2)
    del data_col2[0:2]
    # 获取异常复位次数
    data_col3 = sheet_1.col_values(3)
    del data_col3[0:2]
    # 获取丢包次数
    data_col4 = sheet_1.col_values(4)
    del data_col4[0:2]

    # 读取sheet0 的数据写入Y轴
    sheet = work_book.sheet_by_index(0)
    # 入网时间
    data_sheet0_col1 = str_int(sheet.col_values(1))
    # print(type(data_sheet0_col1[0]))
    # data_sheet0_col1 = np.divide(data_sheet0_col1, 10)
    # 云平台注册——订阅时间
    data_sheet0_col4 = str_int(sheet.col_values(4))
    # 云平台发送数据时间
    data_sheet0_col5 = str_int(sheet.col_values(5))
    # SNR ，RSRP
    data_sheet0_col10 = str_int(sheet.col_values(10))
    # data_sheet0_col10 = np.divide(data_sheet0_col10, 10)
    data_sheet0_col11 = str_int(sheet.col_values(11))
    # print(data_sheet0_col11)
    # data_sheet0_col11 = np.divide(data_sheet0_col11, 10)
    # ping业务
    data_sheet0_col19 = str_int(sheet.col_values(19))
    data_sheet0_col19 = np.divide(data_sheet0_col19, 1000)
    data_sheet0_col20 = str_int(sheet.col_values(20))
    data_sheet0_col20 = np.divide(data_sheet0_col20, 1000)
    data_sheet0_col21 = str_int(sheet.col_values(21))
    data_sheet0_col21 = np.divide(data_sheet0_col21, 1000)
    # print("读取EXCEl时间：" + str(format(time.time() - excel_time, '.2f')))
    my_log.debug("读取EXCEl时间：" + str(format(time.time() - excel_time, '.2f')))
    # col1_3_value = str_int(sheet.col_values(3))
    # del data_sheet0_col3[0:2]


    # X轴数据
    xdata = np.arange(len(data_col4))
    # 使能交互模式
    plt.ion()
    line1 = ax[0].plot(xdata, data_col1, label="get_net_timeout", color=color[0], linewidth=0.5)  # , lable="长时间不入网"
    line2 = ax[0].plot(xdata, data_col2, label="CTwing_Fail", color=color[1], linewidth=0.5)  # , lable="云平台注册"
    line3 = ax[0].plot(xdata, data_col3, label="Abnormal_RST", color=color[2], linewidth=0.5)  # , lable="异常复位"
    line4 = ax[0].plot(xdata, data_col4, label="Ping_loss/10", color=color[3], linewidth=0.5)  # , lable="丢包次数"

    # 画第二个坐标轴
    line5 = ax11.plot(xdata, data_sheet0_col1, label="net_time", color=color[4], linewidth=0.5)  # 入网时间/10
    # line5 = ax[1].plot(xdata, data_sheet0_col1, label="net_time/10", color=color[4], linewidth=0.5)  # 入网时间/10
    line6 = ax[1].plot(xdata, data_sheet0_col4, label="Reg_Sub_time", color=color[5], linewidth=0.5)  # 云平台登录注册总时间
    line7 = ax[1].plot(xdata, data_sheet0_col5, label="send_data_time", color=color[6], linewidth=0.5) # 云平台发送数据时间


    # line9 = ax[2].plot(xdata, data_sheet0_col10, label="RSRP/10", color=color[7], linewidth=0.5)
    ax22.plot(xdata, data_sheet0_col10, label="RSRP", color=color[7], linewidth=0.5)
    line10 = ax[2].plot(xdata, data_sheet0_col11, label="SNR", color=color[8], linewidth=0.5)
    # ping业务
    line11 = ax[3].plot(xdata, data_sheet0_col19, label="Min/1000", color=color[9], linewidth=0.5)
    line12 = ax[3].plot(xdata, data_sheet0_col20, label="Max/1000", color=color[10], linewidth=0.5)
    line13 = ax[3].plot(xdata, data_sheet0_col21, label="Arg/1000", color=color[11], linewidth=0.5)

    # line5 = ax[0].plot(xdata, col1_3_value, label="注册时间", color="pink", linewidth=1)  # , lable="丢包次数"
    title_name = str(time_excel) + '_' + str(test_port) + '_Test Mode_' + str(power) + '_非PSM'
    if j == 0:
        # plt.xlabel("times")
        # plt.ylabel("Abnomal")
        # plt.title(title_name, color='red', loc='center')

        # plt.title(title_name, x=0.1, y=2.3, color='red')
        plt.title(title_name, x=0.5, y=3, color='red')
        # ax[0, 0].set_title(title_name, color='red', loc='center')
        # plt.set_ylabel("Time(s)", loc='center')
        # 入网时间,y云平台注册，订阅时间
        # ax[1].set_ylim(-1, 12)
        # ax11.set_ylim(-1, 120)
        # # ax[1].yticks(-2,0,2,4,6,10,20)
        ax[0].set_xticklabels([])
        ax[1].set_xticklabels([])
        ax[2].set_xticklabels([])
        ax[1].set_yticks([-1, 1.5, 5, 12])
        ax11.set_yticks([-1, 20, 50, 120])
        # a2.set_yticks([-2, 4, 10, 15])
        # ping包最大时延，最小时延
        ax[3].set_ylim(0, 10)
        # RSRP和SNR
        ax[2].set_ylim(-20, 30)
        ax[2].set_yticks([-8, 0, 10, 25])
        ax22.set_yticks([-60, -80, -90, -120, -140])
        # a2.set_ylim(-10, 25)
        # print(plt)
        ax[0].legend(loc='upper right', fontsize=6)
        ax[1].legend(loc='upper left', fontsize=6)
        ax[2].legend(loc='upper left', fontsize=6)
        ax[3].legend(loc='upper right', fontsize=6)
        ax11.legend(loc='upper right', fontsize=6)
        ax22.legend(loc='upper right', fontsize=6)

    # Need both of these in order to rescale
    ax[0].relim()
    ax[0].autoscale_view()
    ax[1].relim()
    ax[1].autoscale_view()
    # draw and flush the figure .
    figure.canvas.draw()
    figure.canvas.flush_events()
    # 关闭交互模式
    # plt.tight_layout()

    if drawline == '1':
        # print
        plt.show()
    plt.ioff()
    plt.savefig(path + '\\' + title_name + '.png', dpi=300, bbox_inches='tight')
    # figsize = (24, 8)
    my_log.debug("读取excel和画图所需时间：" + str(format(time.time() - excel_time, '.2f')))


def restart_test():
    global task_end_time
    t = get_net_time()
    if t == -1:
        list = ['-1', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '9999', '0', '500']
        return list
    time.sleep(1)  # 休眠等待入网成功后打印出的一些数据
    my_log.debug('休眠等待入网成功后打印出的一些数据' + getdata_0())
    list = get_cell_info()
    list.append(t)
    list.append(detach_time)
    # 判断是否做Ping业务
    ping_result = do_ping(ping)
    # 云平台测试
    # CTwing_test(str_CTwing)
    list.append(ping_result)
    # 云平台注册
    CTWing_result = CTWing_case_run(CTwing, CTwing_case_num)
    list.append(CTWing_result[0])
    list.append(CTWing_result[1])
    list.append(CTWing_result[2])
    my_log.debug(list)
    task_end_time = time.time()
    return list


# 采集网络数据
def get_net_data():
    global get_net_data_num
    # print("进入采集数据........")
    if get_net_data_num == '1':
        print("正在采集数据........")
        get_NBSTATS()
        time.sleep(1)
        get_LWIPDATA()


def first_test():
    global list, CT_version, psm_mode, task_end_time, flash, restart_reboot, t1, at_first_time
    # 配置云平台参数
    if CTwing == 1:
        print("配置云平台参数" + str(CTwing_case_num))
        CTWing_case_config(CTwing_case_num, imei)
        if auto_CTwing == "1":
            send_AT0('at+lctregen=1')
        else:
            send_AT0('at+lctregen=0')
        if flash == 0:
            send_AT0(str_REBOOT)
            restart_reboot += 1
    else:
        send_AT0('at+lctregen=0')
    # 是否锁频，在配置参数里面设置，当value值不等于0时，会发送value的相应命令
    if config['learfcn0'] != '0':
        send_AT0(config['learfcn0'])
        print('锁定频点' + config['learfcn0'])
    set_psm_mode()
    if a_p_log == '1':
        tt2 = Thread(target=read_pcore)
        tt2.start()
    send_AT0('AT+REBOOT')
    restart_reboot += 1
    t1 = time.time()

    if psm_mode == "1":
        t = get_net_time()
        at_first_time = t1
        print("第一次入网时间：" + str(t))
        print("下面等待进入休眠...")
        my_log.debug("第一次入网时间：" + str(t))
        my_log.debug("下面等待进入休眠...")
        task_end_time = time.time()
        psm_test()

    else:
        t = get_net_time()
        if t == -1:
            list = ['-1', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0',
                    '0']
            # 把不入网，丢包，云平台注册3个参数写到其他excel
            list1.append(list[0])
            list1.append(list[17])
            list2.append(list[3])
            # 获取网络注册信息
            get_net_data()
            re_start()
            return list
        # time.sleep(1)  # 休眠等待入网成功后打印出的一些数据
        # my_log.debug('休眠等待入网成功后打印出的一些数据' + getdata_0())
        list.append(t)
        list.append(detach_time)
        # 云平台测试
        CTWing_case_run(CTwing, CTwing_case_num)
        # 获取小区信息
        get_cell()
        if config['CT_version'] != '1':
            get_cell_csinfo()
        else:
            list.append("0")
            list.append("0")
            list.append("0")
            list.append("0")
            list.append("0")

        # 做Ping业务
        do_ping(ping)
        my_log.debug(list)
        list1.append(list[0])
        list1.append(list[17])
        list2.append(list[3])
        get_net_data()
    return list


def star_test():
    global list, psm_mode, CT_version, psm_get_net_time, psm_active
    if "at_ldebug" in config:
        if config['at_ldebug'] != '0':
            c = j//int(config['at_ldebug'])
            b = c%4
            if b == 0:
                send_AT0("at+ldebug=0")
                print('-----------------------------------------发送at+ldebug=0------------------------------------')
            elif b == 1:
                send_AT0("at+ldebug=2")
                print('-----------------------------------------发送at+ldebug=2------------------------------------')
            elif b == 2:
                send_AT0("at+ldebug=3")
                print('-----------------------------------------发送at+ldebug=3------------------------------------')
            else:
                send_AT0("at+ldebug=1")
                print('-----------------------------------------发送at+ldebug=1------------------------------------')

    if psm_mode == "1":
        psm_test()
    else:
        if config["demo_test"] == '1':
            list.append(0)
            list.append(0)
            CTwing_case_num = 11
            n = j % 3 + 1
            time.sleep(int(config['auto_send_time']) * n)
        else:
            # n = j % 3 + 1
            # time.sleep(int(config['auto_send_time'])*n)
            # 每次入网切换频点
            if config['change_earfcn_test'] == '1':
                if j % 3 == 0:
                    send_AT0(config['change_earfcn1'])
                    print("锁频点:" + config['change_earfcn1'])
                elif j % 3 == 1:
                    send_AT0(config['change_earfcn2'])
                    print("锁频点:" + config['change_earfcn2'])
                else:
                    send_AT0(config['change_earfcn3'])
                    print("锁频点:" + config['change_earfcn3'])

            if power == '1' or power == '2':
                cfun_0_1()
            if power == '3':
                enble_0_1()
            if power == '4':
                reboot()
            t = get_net_time()
            if t == -1:
                list = ['-1', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '9999', '0', '500']
                # 把不入网，丢包，云平台注册3个参数写到其他excel
                list1.append(list[0])
                list1.append(list[17])
                list2.append(list[3])
                # 获取网络注册信息
                get_net_data()
                re_start()
                return list
            # time.sleep(1)  # 休眠等待入网成功后打印出的一些数据
            # my_log.debug('休眠等待入网成功后打印出的一些数据' + getdata_0())
            list.append(t)
            list.append(detach_time)
        # 云平台测试
        CTWing_case_run(CTwing, int(config["CTwing_case_num"]))
        # 获取小区信息
        get_cell()
        if config['CT_version'] != '1':
            get_cell_csinfo()
        else:
            list.append("0")
            list.append("0")
            list.append("0")
            list.append("0")
            list.append("0")

        # 做Ping业务
        do_ping(ping)
        # print(list)
        my_log.debug(list)
        list1.append(list[0])
        list1.append(list[17])
        list2.append(list[3])
        get_net_data()
    return list


# 设置Excel表头信息
def set_sheet(sheet):
    if sheet == work_sheet:
        work_sheet.write(1, 0, "次数")
        work_sheet.write(1, 1, 'Time')
        work_sheet.write(1, 2, 'Detach_time')
        work_sheet.write(1, 3, 'Reg_time')
        work_sheet.write(1, 4, 'Sub_time')
        work_sheet.write(1, 5, 'Send_data_time')
        work_sheet.write(1, 6, 'Dereg_time')
        work_sheet.write(1, 7, '频点')
        work_sheet.write(1, 8, 'CELLID')
        work_sheet.write(1, 9, 'PCI')
        work_sheet.write(1, 10, 'RSRP')
        work_sheet.write(1, 11, 'SNR')
        work_sheet.write(1, 12, 'Peak1')
        work_sheet.write(1, 13, 'Peak2')
        work_sheet.write(1, 14, 'CellID_old')
        work_sheet.write(1, 15, 'History')
        work_sheet.write(1, 16, 'T0')

        work_sheet.write(1, 17, 'ping_result')
        work_sheet.write(1, 18, 'ping_loss')
        work_sheet.write(1, 19, 'Min')
        work_sheet.write(1, 20, 'Max')
        work_sheet.write(1, 21, 'Arg')

    if sheet == work_sheet1:
        work_sheet1.write(1, 0, "次数")
        work_sheet1.write(1, 1, "入网时间")
        work_sheet1.write(1, 2, "Ping_loss")
        work_sheet1.write(1, 3, "AccessReqcount")
        work_sheet1.write(1, 4, 'RARFailcount')
        work_sheet1.write(1, 5, 'Msg4Failcount')
        work_sheet1.write(1, 6, 'ContentionFailcount')
        work_sheet1.write(1, 7, 'AuthREQcount')
        work_sheet1.write(1, 8, 'AuthFailcount')
        work_sheet1.write(1, 9, 'AuthRejectcount')
        work_sheet1.write(1, 10, 'AttachREQCount')
        work_sheet1.write(1, 11, 'AttachRejectcount')
        work_sheet1.write(1, 12, 'rsrp_arg')
        work_sheet1.write(1, 13, 'rsrp_min')
        work_sheet1.write(1, 14, 'rsrp_max')
        work_sheet1.write(1, 15, 'snr_arg')
        work_sheet1.write(1, 16, 'snr_min')
        work_sheet1.write(1, 17, 'snr_max')
        work_sheet1.write(1, 18, 'TotalAGC_arg')
        work_sheet1.write(1, 19, 'TotalAGC_min')
        work_sheet1.write(1, 20, 'TotalAGC_max')
        work_sheet1.write(1, 21, 'PwrRACH_arg')
        work_sheet1.write(1, 22, 'PwrRACH_min')
        work_sheet1.write(1, 23, 'PwrRACH_max')
        work_sheet1.write(1, 24, 'PwrPusch_arg')
        work_sheet1.write(1, 25, 'PwrPusch_min')
        work_sheet1.write(1, 26, 'PwrPusch_max')
        work_sheet1.write(1, 27, 'csFailCnt')
        work_sheet1.write(1, 28, 'mibFailCnt')
        work_sheet1.write(1, 29, 'sib1FailCnt')
        work_sheet1.write(1, 30, 'siFailCnt')
        work_sheet1.write(1, 31, 'pdschDecodeFailCnt')
        work_sheet1.write(1, 32, 'pdcchDecodeFailCnt')
        work_sheet1.write(1, 33, 'csUseRegInfoCnt')
        work_sheet1.write(1, 34, 'csNotUseRegInfoCnt')
        work_sheet1.write(1, 35, 'retransmitCnt')

    if sheet == work_sheet2:
        work_sheet2.write(1, 0, '次数')
        work_sheet2.write(1, 1, '云平台注册')
        work_sheet2.write(1, 2, 'TxPacketCnt')
        work_sheet2.write(1, 3, 'RxPacketCnt')
        work_sheet2.write(1, 4, 'TxLossPacketCnt')
        work_sheet2.write(1, 5, 'TxMaxPacketSize')
        work_sheet2.write(1, 6, 'RxMaxPacketSize')
        work_sheet2.write(1, 7, 'TxPacketByte')
        work_sheet2.write(1, 8, 'RxPacketByte')

    if sheet == work_sheet3:
        work_sheet3.write(1, 0, '次数')
        work_sheet3.write(1, 1, '长时间不入网')
        work_sheet3.write(1, 2, '云平台注册失败')
        work_sheet3.write(1, 3, '异常复位统计')
        work_sheet3.write(1, 4, 'Ping丢包统计')
        # work_sheet3.write(1, 5, 'RxMaxPacketSize')
        # work_sheet3.write(1, 6, 'TxPacketByte')
        # work_sheet3.write(1, 7, 'RxPacketByte')

    if sheet == work_sheet4:
        work_sheet4.write(1, 0, '次数')
        work_sheet4.write(1, 1, 'all_time')
        work_sheet4.write(1, 2, '业务结束-进入深睡时间')
        work_sheet4.write(1, 3, 'PSM入网时间')
        work_sheet4.write(1, 4, '频点')
        work_sheet4.write(1, 5, 'CELLID')
        work_sheet4.write(1, 6, 'PCI')
        work_sheet4.write(1, 7, 'RSRP')
        work_sheet4.write(1, 8, 'SNR')
        work_sheet4.write(1, 9, '云平台登录')
        work_sheet4.write(1, 10, '云平台注册')
        work_sheet4.write(1, 11, '发送数据')
        work_sheet4.write(1, 12, '注销云平台')
        work_sheet4.write(1, 13, '唤醒-业务-休眠时间')


if __name__ == '__main__':
    global j
    # 新建工作薄
    work_book = xlwt.Workbook()
    # 获取sheet
    work_sheet = work_book.add_sheet('测试数据')
    work_sheet1 = work_book.add_sheet('网络监控信息')
    work_sheet2 = work_book.add_sheet('云平台监控信息')
    work_sheet3 = work_book.add_sheet('测试结果汇总')
    work_sheet4 = work_book.add_sheet('PSM模式测试')
    set_sheet(work_sheet)
    set_sheet(work_sheet1)
    set_sheet(work_sheet2)
    set_sheet(work_sheet3)
    set_sheet(work_sheet4)
    excel_name = 'Test' + time_excel + test_port + '.csv'
    work_book.save(path + '\\' + excel_name)
    if a_p_log == '2':
        # ser.close()
        tt1 = Thread(target=read_acore)
        tt1.start()
        tt2 = Thread(target=read_pcore)
        tt2.start()
        while 1:
            time.sleep(1000000)
            pass

    if a_p_log == '1':
        tt1 = Thread(target=read_acore)
        tt1.start()
        # tt2 = Thread(target=read_pcore)
        # tt2.start()2

    if random_power == '1':
        tt3 = Thread(target=random_power1)
        tt3.start()
    # 发送一个AT防止psm模式默认打开
    # send_AT0('AT')
    set_psm_mode_0()
    # time.sleep(0.5)
    getdata_0()
    # 查看软件版本
    if config['CT_version'] != '1':
        send_AT0("at+lver")
        time.sleep(1.5)
        version_data = getdata().strip().rstrip("OK").strip().lstrip("+LVER: ")

    print('############################ 启动NB-IoT 终端入网、退网、ping业务、PSM、CT云平台压力测试###########################')
    if config['CT_version'] != '1':
        print("# 固件版本：" + version_data)
        work_sheet.write(0, 0, version_data)
    print('# 硬件版本: LZ8001E_HW_1.0')
    print("# 测试时间：" + str(time_excel))
    print('#################################@广州粒子微电子有限公司020-82510621###############################')
    if psm_mode != '1':
        power_mode = "电源模式：" + str(power) + '\n' + "at+lsdata状态：" + str(at_lsdata) + '\n' + "测试的端口号：" + str(test_port) + '\n' + "是否初始化：" + str(init_test)
        work_sheet.write(0, 1, power_mode)

    imei = get_imei()
    times = 1000000
    j = 0
    ping_min = 0
    ping_max = 40000
    while j < times:
        try:
            print("-------------------------------------第" + str(j + 1) + "次测试开始------------------------------------------------")

            if init_test == 1 and j == 0:
                print("初始化测试")
                init()
                # 回读acore是否开启flash保护
                if config['a_p_log'] != '0':
                    if config['flash_p'] == '1':
                        print("发送cfun0，开启flash保护")
                        my_log.debug("发送cfun0，开启flash保护")
                        getdata_sleep(2)
                        ser_acore.write(bytes("flash p" + "\r\n", encoding="utf-8"))
                        getdata_sleep(2)
                    else:
                        print("关闭flash保护")
                        my_log.debug("关闭flash保护")
                        getdata_sleep(2)
                        ser_acore.write(bytes("flash u" + "\r\n", encoding="utf-8"))
                        getdata_sleep(2)
                else:
                    print('未开启acore')
                    my_log.debug('未开启acore')
                my_log.debug(config)
            elif init_test != 1 and j == 0:
                first_test()
                if config['a_p_log'] != '0':
                    if config['flash_p'] == '1':
                        print("发送cfun0，开启flash保护")
                        my_log.debug("发送cfun0，开启flash保护")
                        getdata_sleep(2)
                        ser_acore.write(bytes("flash p" + "\r\n", encoding="utf-8"))
                        getdata_sleep(2)
                    else:
                        print("关闭flash保护")
                        my_log.debug("关闭flash保护")
                        getdata_sleep(2)
                        ser_acore.write(bytes("flash u" + "\r\n", encoding="utf-8"))
                        getdata_sleep(2)
                else:
                    print('未开启acore')
                    my_log.debug('未开启acore')
                my_log.debug(config)

            else:
                star_test()

            if psm_mode == "1":
                print(list3)
                # my_log.debug(list3)
                # 异常复位统计
                RST_total = RST_0 + RST_1 + RST_2 + RST_3 + RST_4 + RST_6 - restart_reboot - restart_power
                # print("发AT激活次数：" + str(restart_psm))
                my_log.debug("psm复位减去次数：" + str(restart_psm) + "reboot次数：" + str(restart_reboot) + "power次数：" + str(restart_power))
                my_log.debug("发AT激活次数：" + str(restart_psm))
                my_log.debug("总共PSM复位次数：" + str(RST_5))

                if RST_total != 0:
                    d = RST_total
                else:
                    d = j + 1
                RST_0_rate = ",(" + str(format(((RST_0 - restart_power) / d) * 100, '.2f')) + "%)"
                RST_1_rate = ",(" + str(format((RST_1 / d) * 100, '.2f')) + "%)"
                RST_2_rate = ",(" + str(format(((RST_2 - restart_reboot) / d) * 100, '.2f')) + "%)"
                RST_3_rate = ",(" + str(format((RST_3 / d) * 100, '.2f')) + "%)"
                RST_4_rate = ",(" + str(format((RST_4 / d) * 100, '.2f')) + "%)"
                # RST_5_rate = ",(" + str(format((RST_5 / d) * 100, '.2f')) + "%)"
                RST_6_rate = ",(" + str(format((RST_6 / d) * 100, '.2f')) + "%)"
                RST_total_rate = ",(" + str(format((RST_total / (j + 1)) * 100, '.2f')) + "%)"

                CT_rate_reg = ",(" + str(format((CTwing_reg / (j + 1)) * 100, '.2f')) + "%)"
                CT_rate_sub = ",(" + str(format((CTwing_sub / (j + 1)) * 100, '.2f')) + "%)"
                CT_data_rate = ",(" + str(format((CTwing_data / (j + 1)) * 100, '.2f')) + "%)"
                # psm异常占比
                psm_active_rate = ",(" + str(format((psm_active / (j + 1)) * 100, '.2f')) + "%)"
                psm_net_time_rate = ",(" + str(format((psm_get_net_time / (j + 1)) * 100, '.2f')) + "%)"

                print("PSM进入休眠失败：" + str(psm_active) + str(psm_active_rate) +
                      "  psm入网失败：" + str(psm_get_net_time) + str(psm_net_time_rate))

                my_log.error("PSM进入休眠失败：" + str(psm_active) + str(psm_active_rate) +
                      "  psm入网失败：" + str(psm_get_net_time) + str(psm_net_time_rate))

                print("云平台注册失败：" + str(CTwing_reg) + str(CT_rate_reg) +
                      "  订阅失败：" + str(CTwing_sub) + str(CT_rate_sub) +
                      "  发送数据失败：" + str(CTwing_data) + str(CT_data_rate))

                my_log.error("云平台注册失败：" + str(CTwing_reg) + str(CT_rate_reg) +
                             "  订阅失败：" + str(CTwing_sub) + str(CT_rate_sub) +
                             "  发送数据失败：" + str(CTwing_data) + str(CT_data_rate))

                print("异常复位统计:" + str(RST_total) + RST_total_rate +
                      "  HW_RST:" + str(RST_0 - restart_power) + str(RST_0_rate) +
                      '  WDG_RST: ' + str(RST_1) + str(RST_1_rate) +
                      '  AT_RST: ' + str(RST_2 - restart_reboot) + str(RST_2_rate) +
                      '  PS_RST: ' + str(RST_3) + str(RST_3_rate) +
                      '  OTA_RST: ' + str(RST_4) + str(RST_4_rate) +
                      # '  PSM_RST: ' + str(RST_5) + str(RST_5_rate) +
                      '  CFUN0_TIMEOUT: ' + str(RST_6) + str(RST_6_rate))

                my_log.error("异常复位统计:" + str(RST_total) + RST_total_rate +
                             "  HW_RST:" + str(RST_0 - restart_power) + str(RST_0_rate) +
                             '  WDG_RST: ' + str(RST_1) + str(RST_1_rate) +
                             '  AT_RST: ' + str(RST_2 - restart_reboot) + str(RST_2_rate) +
                             '  PS_RST: ' + str(RST_3) + str(RST_3_rate) +
                             '  OTA_RST: ' + str(RST_4) + str(RST_4_rate) +
                             # '  PSM_RST: ' + str(RST_5) + str(RST_5_rate) +
                             '  CFUN0_TIMEOUT: ' + str(RST_6) + str(RST_6_rate))
                list_result.append(RST_total)
            else:
                print(list)
                t_rate = ",(" + str(format((long_time / (j + 1)) * 100, '.2f')) + "%)"
                at_rate = ",(" + str(format((AT_die / (j + 1)) * 100, '.2f')) + "%)"
                cell_rate = ",(" + str(format((CELLID_error / (j + 1)) * 100, '.2f')) + "%)"
                at_death_rate = ",(" + str(format((AT_death / (j + 1)) * 100, '.2f')) + "%)"
                # 云平台注册统计

                if list[5] == '注销失败':
                    CTwing_deteach += 1

                CT_rate_reg = ",(" + str(format((CTwing_reg/(j + 1)) * 100, '.2f')) + "%)"
                CT_rate_sub = ",(" + str(format((CTwing_sub/(j + 1)) * 100, '.2f')) + "%)"
                CT_data_rate = ",(" + str(format((CTwing_data/(j + 1)) * 100, '.2f')) + "%)"
                CTwing_deteach_rate = ",(" + str(format((CTwing_deteach/(j + 1)) * 100, '.2f')) + "%)"

                print("长时间不入网：" + str(long_time) + str(t_rate) +
                      "  AT响应异常：" + str(AT_die) + str(at_rate) +
                      "  AT无响应：" + str(AT_death) + str(at_death_rate) +
                      "  选择到不同小区： " + str(CELLID_error) + str(cell_rate))
                print("云平台注册失败：" + str(CTwing_reg) + str(CT_rate_reg) +
                      "  订阅失败：" + str(CTwing_sub) + str(CT_rate_sub) +
                      "  发送数据失败：" + str(CTwing_data) + str(CT_data_rate) +
                      "  注销失败：" + str(CTwing_deteach) + str(CTwing_deteach_rate))

                my_log.error("长时间不入网：" + str(long_time) + str(t_rate) +
                             "  AT响应异常：" + str(AT_die) + str(at_rate) +
                             "  AT无响应：" + str(AT_death) + str(at_death_rate) +
                             "  选择到不同小区： " + str(CELLID_error) + str(cell_rate))
                my_log.error("云平台注册失败：" + str(CTwing_reg) + str(CT_rate_reg) +
                             "  订阅失败：" + str(CTwing_sub) + str(CT_rate_sub) +
                             "  发送数据失败：" + str(CTwing_data) + str(CT_data_rate))

                # 异常复位统计
                RST_total = RST_0 + RST_1 + RST_2 + RST_3 + RST_4 + RST_5 + RST_6 - restart_reboot - restart_power - restart_psm
                my_log.debug("psm复位减去次数：" + str(restart_psm) + "reboot次数：" + str(restart_reboot) + "power次数：" + str(restart_power))
                if RST_total != 0:
                    d = RST_total
                else:
                    d = j + 1
                RST_0_rate = ",(" + str(format(((RST_0 - restart_power)/d) * 100, '.2f')) + "%)"
                RST_1_rate = ",(" + str(format((RST_1/d) * 100, '.2f')) + "%)"
                RST_2_rate = ",(" + str(format(((RST_2 - restart_reboot) / d) * 100, '.2f')) + "%)"
                RST_3_rate = ",(" + str(format((RST_3 / d) * 100, '.2f')) + "%)"
                RST_4_rate = ",(" + str(format((RST_4 / d) * 100, '.2f')) + "%)"
                RST_5_rate = ",(" + str(format((RST_5 / d) * 100, '.2f')) + "%)"
                RST_6_rate = ",(" + str(format((RST_6 / d) * 100, '.2f')) + "%)"
                RST_total_rate = ",(" + str(format((RST_total / (j + 1)) * 100, '.2f')) + "%)"

                print("异常复位统计:" + str(RST_total) + RST_total_rate +
                      "  HW_RST:" + str(RST_0 - restart_power) + str(RST_0_rate) +
                      '  WDG_RST: ' + str(RST_1) + str(RST_1_rate) +
                      '  AT_RST: ' + str(RST_2 - restart_reboot) + str(RST_2_rate) +
                      '  PS_RST: ' + str(RST_3) + str(RST_3_rate) +
                      '  OTA_RST: ' + str(RST_4) + str(RST_4_rate) +
                      '  PSM_RST: ' + str(RST_5) + str(RST_5_rate) +
                      '  CFUN0_TIMEOUT: ' + str(RST_6) + str(RST_6_rate))

                my_log.error("异常复位统计:" + str(RST_total) + RST_total_rate +
                             "  HW_RST:" + str(RST_0 - restart_power) + str(RST_0_rate) +
                             '  WDG_RST: ' + str(RST_1) + str(RST_1_rate) +
                             '  AT_RST: ' + str(RST_2 - restart_reboot) + str(RST_2_rate) +
                             '  PS_RST: ' + str(RST_3) + str(RST_3_rate) +
                             '  OTA_RST: ' + str(RST_4) + str(RST_4_rate) +
                             '  PSM_RST: ' + str(RST_5) + str(RST_5_rate) +
                             '  CFUN0_TIMEOUT: ' + str(RST_6) + str(RST_6_rate))
                list_result.append(long_time)
                list_result.append(CTwing_reg)
                list_result.append(RST_total)
            if ping == 1:
                # ping业务统计
                ping_a = ping_1 + ping_2
                ping_a_t = int(ping_a)/10
                list_result.append(ping_a_t)
                if j == 0:
                    ping_min = int(list[18])
                    ping_max = int(list[19])
                    ping_avg = int(list[20])
                if ping_min >= int(list[18]):
                    if ping_min == 0:
                        ping_min = 5000
                    else:
                        ping_min = int(list[18])
                if ping_max <= int(list[19]):
                    ping_max = int(list[19])
                # 获取到上一次的平均时延
                # ping_avg2 = int(list[16])
                # 计算实际发起ping包的总数
                ping_true = (j + 1 - long_time - CEREG_0)*ping_total_all
                # print(CEREG_0)
                ping_avg1 = int(list[20])*10*(j+1)
                ping_avg = format((int(ping_avg)*10 + ping_avg1)/(10*(j+2)), '.0f')
                if ping_true == 0:
                    PING_a_rate = ", (0%)"
                else:
                    PING_a_rate = ",(" + str(format(((ping_a)/int(ping_true)) * 100, '.2f')) + "%)"
                print("ping总丢包数：" + str(ping_a) + PING_a_rate + '  +LPING:1: ' + str(ping_1) + '  +LPING:2: ' + str(ping_2) +
                      "  最小时延：" + str(ping_min) + "  最大时延：" + str(ping_max) + "  平均时延：" + str(ping_avg))

                my_log.error("ping总丢包数：" + str(ping_a) + PING_a_rate + '  +LPING:1: ' + str(ping_1) + '  +LPING:2: ' + str(ping_2) +
                      "  最小时延：" + str(ping_min) + "  最大时延：" + str(ping_max) + "  平均时延：" + str(ping_avg))

        except Exception as e:
            print(e.__str__() + "\n出现异常，等待10s，进行下一轮测试")
            time.sleep(2)
            my_log.debug("出现异常读取数据：" + str(ser.read_all()))
            # send_AT0(str_REBOOT)
            # re_start()
            time.sleep(8)
            my_log.debug("出现异常读取数据：" + str(ser.read_all()))
            while len(list) < 21:
                list.append('0')
            while len(list3) < 10:
                list3.append('0')
        draw_line_time1 = time.time()
        time_cur = time.strftime('%m-%d-%H-%M-%S', time.localtime(time.time()))
        work_sheet.write(j + 2, 0, "第" + str(j + 1) + "次测试" + time_cur)
        work_sheet1.write(j + 2, 0, "第" + str(j + 1) + "次测试" + time_cur)
        work_sheet2.write(j + 2, 0, "第" + str(j + 1) + "次测试" + time_cur)
        work_sheet3.write(j + 2, 0, "第" + str(j + 1) + "次测试" + time_cur)
        work_sheet4.write(j + 2, 0, "第" + str(j + 1) + "次测试" + time_cur)
        if psm_mode == '1':
            my_log.error("第" + str(j + 1) + "次：" + str(list3))
        else:
            my_log.error("第" + str(j + 1) + "次：" + str(list))
        time.sleep(2)
        for i in range(len(list)):
            work_sheet.write(j + 2, i + 1, list[i])
        for i in range(len(list1)):
            work_sheet1.write(j + 2, i + 1, list1[i])
        for i in range(len(list2)):
            work_sheet2.write(j + 2, i + 1, list2[i])
        for i in range(len(list_result)):
            work_sheet3.write(j + 2, i + 1, list_result[i])
        for i in range(len(list3)):
            work_sheet4.write(j + 2, i + 1, list3[i])

        # 保存Excel
        work_book.save(path + '\\' + excel_name)
        # os.chmod(path + '\\' + excel_name, 0o444)
        # draw_line_time1 = time.time()
        if psm_mode == '1':
            # y轴业务结束-唤醒
            list_psm_active_y.append(psm_active)
            # y轴不入网次数
            list_psm_get_net_y.append(psm_get_net_time)

        if j%50 == 0:
            try:
                if psm_mode != '1':
                    draw_line(excel_name)
                else:
                    draw_line_psm(excel_name)
            except Exception as e:
                print(e.__str__() + "\n出现异常s，进行下一轮测试")
        draw_line_time = format(time.time()-draw_line_time1, '.2f')
        # print("输出request次数：" + str(request_data) + "  输出reject次数：" + str(reject))
        # my_log.error("输出request次数：" + str(request_data) + "  输出reject次数：" + str(reject))
        time.sleep(sleep_time)
        my_log.error(config_acore)
        print(config_acore)
        my_log.error(config_pcore)
        print(config_pcore)
        print("USIM: " + str(dic_sim))
        my_log.error("USIM: " + str(dic_sim))
        print("flash_change:" + str(flash_change))
        my_log.error("flash_change:" + str(flash_change))
        print("小区变化后RSRP变化-6dbm：" + str(rsrp_change_times))
        my_log.error("小区变化后RSRP变化-6dbm：" + str(rsrp_change_times))

        list.clear()
        list1.clear()
        list2.clear()
        list_result.clear()
        list3.clear()
        j += 1
        if 'waite_time' in config:
            waite_time = int(config['waite_time'])
            # print(config['waite_time'])
            print("等待" + str(waite_time) + "s下一次开始.......")
            getdata_sleep(waite_time)
            # time.sleep(waite_time)
        # time.sleep(40)
