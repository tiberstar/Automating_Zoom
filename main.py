import os
import sys
import pandas as pd
import pyautogui
import logging
import urllib3
import time
from datetime import datetime

# 入会结果是否发送微信消息
sendResult = True
# 微信绑定SDK
sdk = 'SCT97191T10sJtQMJbP8e02jPQbJg5iCs'
# 腾讯会议应用程序路径
meetPath = "D:/Program Files (x86)/Tencent/Meet/WeMeet/wemeetapp.exe"
# 微信推送消息标题
successTitle = "入会成功!"
failTitle = "入会失败!"

# logging.basicConfig(filename="server.log",filemode="w",format="[%(asctime)s]-[%(name)s]-[%(levelname)s] %(message)s",level=logging.INFO)
# 创建一个logger
logger = logging.getLogger('mylogger')
logger.setLevel(logging.DEBUG)
# 创建一个handler，用于写入日志文件
fh = logging.FileHandler('server.log')
fh.setLevel(logging.DEBUG)
# 再创建一个handler，用于输出到控制台
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)
# 定义handler的输出格式
formatter = logging.Formatter("[%(asctime)s][%(thread)d][%(filename)s][line: %(lineno)d][%(levelname)s] ## %(message)s")
fh.setFormatter(formatter)
ch.setFormatter(formatter)
# 给logger添加handler
logger.addHandler(fh)
logger.addHandler(ch)

# 需要一个PoolManager实例来生成请求,由该实例对象处理与线程池的连接以及线程安全的所有细节，不需要任何人为操作：
http = urllib3.PoolManager()
global meetingState
meetingState = False

def getPath(filename):
    # 方法一（如果要将资源文件打包到app中，使用此法）
    bundle_dir = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    path = os.path.join(bundle_dir, filename)
    return path

# Excel文件路径
excelPath = getPath("meetingList.xlsx")
df = pd.read_excel(excelPath)

def signIn(meeting_id, password):
    # 从指定位置打开应用程序
    os.startfile(meetPath)
    time.sleep(5)

    # 点击加入按钮
    joinbtn = pyautogui.locateCenterOnScreen(getPath("buttons/joinameeting.png"))
    pyautogui.moveTo(joinbtn)
    pyautogui.click()
    time.sleep(2)

    # 输入会议 ID
    try:
        meetingidbtn = pyautogui.locateCenterOnScreen(getPath("buttons/meetingId.png"))
        pyautogui.moveTo(meetingidbtn)
        pyautogui.write(meeting_id)
        time.sleep(2)
    except:
        meetingidbtn = pyautogui.locateCenterOnScreen(getPath("buttons/inputMeetingId.png"))
        pyautogui.moveTo(meetingidbtn)
        pyautogui.write(meeting_id)
        time.sleep(2)

    # 关闭视频和音频
    try:
        mediaBtn = pyautogui.locateAllOnScreen(getPath("buttons/media.PNG"))
        pyautogui.moveTo(mediaBtn)
        pyautogui.click()
    except:
        logger.info("入会开启摄像头已关闭，无需自动取消。")
    try:
        audioBtn = pyautogui.locateAllOnScreen(getPath("buttons/audio.PNG"))
        pyautogui.moveTo(audioBtn)
        pyautogui.click()
        time.sleep(2)
    except:
        logger.info("入会开启麦克风已关闭，无需自动取消。")

    # 加入
    join = pyautogui.locateCenterOnScreen(getPath("buttons/join.PNG"))
    pyautogui.moveTo(join)
    pyautogui.click()
    time.sleep(2)

    # 输入密码以加入会议
    passcode = pyautogui.locateCenterOnScreen(getPath("buttons/meetingPasscode.PNG"))
    pyautogui.moveTo(passcode)
    pyautogui.write(password)
    time.sleep(1)

    # 点击加入按钮
    joinmeeting = pyautogui.locateCenterOnScreen(getPath("buttons/joinmeeting.PNG"))
    pyautogui.moveTo(joinmeeting)
    pyautogui.click()
    time.sleep(2)

def signInExcelMeeting():
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    if now in str(df['Timings']):
        mylist = df["Timings"]
        mylist = [i.strftime("%Y-%m-%d %H:%M") for i in mylist]
        c = [i for i in range(len(mylist)) if mylist[i] == now]
        row = df.loc[c]
        meeting_id = str(row.iloc[0, 1])
        password = str(row.iloc[0, 2])
        global meetingTitle, meetingTime, meetingId
        meetingTitle = str(row.iloc[0, 3])
        meetingTime = str(row.iloc[0, 0]).replace(" ", "%20")
        meetingId = str(row.iloc[0, 1])
        time.sleep(5)
        try:
            signIn(meeting_id, password)
            time.sleep(2)
            logger.info(meeting_id + ":" + meetingTitle + " " + successTitle)
            meetingState = True
            if sendResult == True:
                logger.info("开始推送微信状态...")
                try:
                    url = 'https://sc.ftqq.com/' + sdk + \
                          ".send?title=" + meetingTitle + "%20%20" + successTitle + \
                          "&desp=会议信息如下:" \
                          "%0a%0d会议时间:" + meetingTime + \
                          "%0a%0d会议号:" + meetingId + \
                          "%0a%0d会议主题:" + meetingTitle
                    # 通过request()方法创建一个请求，该方法返回一个HTTPResponse对象：
                    r = http.request('GET', url)
                    logger.info("微信消息发送成功。")
                except Exception as e:
                    logger.info("微信推送失败。")
                    logger.exception(e.args)
            else:
                logger.info("配置为不推送微信状态。")
        except:
            logger.info(meeting_id + failTitle)
            if sendResult == True:
                logger.info("开始推送微信状态...")
                try:
                    url = "https://sc.ftqq.com/" + sdk + \
                          ".send?title=" + meetingTitle + "%20%20" + failTitle + \
                          "&desp=会议信息如下:" \
                          "%0a%0d会议时间:" + meetingTime + \
                          "%0a%0d会议号:" + meetingId + \
                          "%0a%0d会议主题:" + meetingTitle
                    logger.info(url)
                    # 通过request()方法创建一个请求，该方法返回一个HTTPResponse对象：
                    r = http.request('GET', url)
                    logger.info("微信消息发送成功。")
                except Exception as e:
                    logger.info("微信推送失败。")
                    logger.exception(e.args)
            else:
                logger.info("配置为不推送微信状态。")


def quitMeeting():
    try:
        quitBtn = pyautogui.locateCenterOnScreen(getPath("buttons/quitMeeting.png"))
        pyautogui.moveTo(quitBtn)
        pyautogui.click()
    except:
        logger.error("未找到离开会议按钮。")
    time.sleep(2)

    if sendResult:
        logger.info("开始推送微信状态...")
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S").replace(" ", "%20")
        try:
            url = "https://sc.ftqq.com/" + sdk + \
                  ".send?title=" + meetingTitle + "%20%20已结束" + \
                  "&desp=会议信息如下:" \
                  "%0a%0d会议时间:" + meetingTime + \
                  "%0a%0d会议号:" + meetingId + \
                  "%0a%0d会议主题:" + meetingTitle + \
                  "%0a%0d会议结束时间:" + now
            # 通过request()方法创建一个请求，该方法返回一个HTTPResponse对象：
            r = http.request('GET', url)
            logger.info("微信消息发送成功。")
        except Exception as e:
            logger.info("微信推送失败。")
            logger.exception(e.args)
    else:
        logger.info("配置为不推送微信状态。")


while True:
    # To get current time
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    logger.info("当前时间:" + now)
    signInExcelMeeting()
    quitBtn = pyautogui.locateCenterOnScreen(getPath("buttons/quitMeeting.png"))
    userSize = pyautogui.locateCenterOnScreen(getPath("buttons/userSize.png"))
    thanks = pyautogui.locateCenterOnScreen(getPath("buttons/thanks.png"))
    # 存在退出按钮和人数37 或者 存在退出按钮和感谢话语
    if (quitBtn != None and userSize != None and meetingState) or (quitBtn != None and thanks != None and meetingState):
        quitMeeting()
    else:
        if meetingState:
            logger.info("会议未结束。")
        elif meetingState == False:
            logger.info("无正在进行中的会议。")

    time.sleep(60)
