# 腾讯会议自动化入会脚本
根据您的时间表自动加入缩放会议的 python 脚本。

<ol>
<li>检查“timings.xlsx”文件以查找即将开始的会议。</li>
<li>一旦当前时间与任何会议时间匹配，它就会打开 Zoom 桌面应用程序。</li>
<li>自动将光标导航到加入会议的各个步骤。</li>
<li>会议 ID 和密码从“meetingList.xlsx”中提取并自动输入到 Zoom 应用程序中。</li>
</ol>

## Prerequisites

<ol>
<li>您的系统中必须安装 Zoom 应用程序且已登录账号。</li>
<li>当天的会议时间以及会议 ID 和密码必须手动输入“meetingList.xlsx”</li>
  <li>pyautogui, python, pandas</li>
</ol>

## Behind the scenes

<ol>
<li>无限循环使用“datetime.now”函数不断检查系统的当前时间。</li>
<li>只要当前时间与“meetingList.xlsx”中提到的时间匹配，就会使用“os.startfile()”函数打开缩放应用程序。</li>
<li>“pyautogui.locateOnScreen()”函数在屏幕上定位加入按钮的图像并返回位置。</li>
<li>“pyautogui.moveTo()”将光标移动到该位置。</li>
<li>“pyautogui.click()”执行点击操作。</li>
<li>使用“pyautogui.write()”命令输入会议 ID 和密码。</li>
</ol>

## License & Copyright

© 2020 <b>Sunil Aleti</b><br>
Licensed under <a href="https://github.com/tiberstar/Automating_Zoom/blob/master/LICENSE">MIT License</a>
