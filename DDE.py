import win32ui
import dde
import time
import os

lst = ["0", "0", "0", "0", "0", "0", "0", "0", "0"]

# 创建一个DDE客户端
dde_client = dde.CreateServer()

# 开始DDE会话
dde_client.Create("MyDDEClient")

# 连接到DDE服务
conversation = dde.CreateConversation(dde_client)

# 连接到特定的服务和主题
while True:
    try:
        conversation.ConnectTo("Orbitron", "Tracking")
        break
    finally:
        print("未检测到Orbitron进程，请确保Orbitron已经启动并打开DDE服务")
        input("请确保Orbitron已经启动并打开DDE服务，短按Enter键进行重新连接")

while True:
    # 请求数据
    data = conversation.Request("TrackingData")
    data = data.strip() if data else data
    if data == "" or data is None:
        print("没有收到DDE数据，请确保Orbitron已经启动并打开DDE服务")
        input("请确保Orbitron已经启动并打开DDE服务，短按Enter键进行重新连接")
    else:
        os.system('cls')
        lst = data.split()
        # 打印接收到的数据
        try:
            print("卫星名称：", lst[0].lstrip("SN"))
        except IndexError:
            print("卫星名称：", "未知")
        try:
            print("卫星方位角：", lst[1].lstrip("AZ"))
        except IndexError:
            print("卫星方位角：", "未知")
        try:
            print("卫星仰角：", lst[2].lstrip("EL"))
        except IndexError:
            print("卫星仰角：", "未知")
        try:
            print("下行多普勒频率：", lst[3].lstrip("DN"))
        except IndexError:
            print("下行多普勒频率：", "未知")
        try:
            print("上行多普勒频率：", lst[4].lstrip("UP"))
        except IndexError:
            print("上行多普勒频率：", "未知")
        try:
            print("下行解调模式：", lst[5].lstrip("MDN"))
        except IndexError:
            print("下行解调模式：", "未知")
        try:
            print("上行调制模式：", lst[6].lstrip("MUP"))
        except IndexError:
            print("上行调制模式：", "未知")
        # 等待1秒
        time.sleep(1)
