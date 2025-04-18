import base64
import hashlib
import time
import requests
from PIL import ImageGrab
import os
import win32com.client
import pygetwindow as gw
import pythoncom
import sys
import io

# 设置 stdout 和 stderr 编码为 utf-8 111
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
if os.name == 'nt':
    os.system('chcp 65001')

# 指定文件路径和对应的 webhook URL
files_and_urls = [
    (r'C:\RPAData\output-401.xlsx',
     'https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=b59764ec-ae6e-4d70-a076-2ec634b1f2f5'),
    #添加第二组文件路径&WebHook地址
    #(r'C:\RPAData\Test-xy-111.xlsx',
     #'https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=ae7ce51a-24f0-4db4-a1be-88d6eaad8883')
]

def log_error(message):
    """记录错误信息到日志文件"""
    with open('error_log.txt', 'a', encoding='utf-8') as log_file:
        log_file.write(message + '\n')

def wait_for_excel_ready(excel_app, timeout=30):
    """等待 Excel 完全加载并准备就绪"""
    start_time = time.time()
    while True:
        if excel_app.Application.Ready:
            break
        if time.time() - start_time > timeout:
            raise TimeoutError("Excel在规定时间内未准备好。")
        time.sleep(0.5)  # 每 500 毫秒检查一次

def bring_excel_to_front():
    """将 Excel 窗口置于前台"""
    windows = gw.getWindowsWithTitle('Excel')
    if windows:
        excel_window = windows[0]
        excel_window.activate()  # 激活 Excel 窗口

def take_screenshots(screenshot_areas):
    """截取多个区域的截图"""
    screenshot_files = []
    for i, area in enumerate(screenshot_areas):
        time.sleep(1)  # 等待 1 秒确保 Excel 完全可见
        screenshot = ImageGrab.grab(bbox=area)
        screenshot_file = f"screenshot_{i + 1}.png"
        screenshot.save(screenshot_file)
        screenshot_files.append(screenshot_file)
    return screenshot_files

def send_image(filename, webhook_url):
    """发送图片到企业微信群"""
    with open(filename, "rb") as f:
        image_data = base64.b64encode(f.read()).decode('utf-8')
        f.seek(0)
        md5_hash = hashlib.md5(f.read()).hexdigest()

    body = {
        'msgtype': 'image',
        'image': {
            'base64': image_data,
            'md5': md5_hash
        }
    }

    # 发送请求到企业微信群
    headers = {'Content-Type': 'application/json;charset=utf-8'}
    response = requests.post(webhook_url, json=body, headers=headers)
    print(response.text)

def process_excel_file(excel_app, file_path, webhook_url):
    """处理单个 Excel 文件：打开、截图、发送"""
    workbook = None
    screenshots = []  # 初始化为空列表

    try:
        # 打开工作簿
        workbook = excel_app.Workbooks.Open(file_path, ReadOnly=True, IgnoreReadOnlyRecommended=True)

        # 等待 Excel 加载完成
        wait_for_excel_ready(excel_app)

        # 将 Excel 窗口置于前台
        bring_excel_to_front()
        time.sleep(3)  # 等待 Excel 完全进入前台

        # 截图区域设置
        screenshot_areas = [
            (50, 295, 740, 895),  # 区域1
            (835, 295, 1660, 900)  # 区域2
        ]

        screenshots = take_screenshots(screenshot_areas)

        for file in screenshots:
            send_image(file, webhook_url)

    except Exception as e:
        print(f"发生错误: {e}")
        log_error(f"发生错误: {e}")

    finally:
        if workbook:
            try:
                # 尝试关闭工作簿
                workbook.Close(SaveChanges=False)
                print("工作簿已成功关闭.")
            except Exception as ce:
                print(f"关闭工作簿时发生错误: {ce}")
                log_error(f"关闭工作簿时发生错误: {ce}")

        # 删除临时文件
        for file in screenshots:
            if os.path.exists(file):
                os.remove(file)
        print("临时文件已删除。")

if __name__ == "__main__":
    pythoncom.CoInitialize()
    excel_app = None

    try:
        # 启动 Excel 应用
        excel_app = win32com.client.Dispatch('Excel.Application')
        excel_app.Visible = True  # 设置 Excel 可见

        for file_path, webhook_url in files_and_urls:
            process_excel_file(excel_app, file_path, webhook_url)

    except Exception as e:
        print(f"发生错误: {e}")
        log_error(f"发生错误: {e}")

    finally:
        # 在所有处理完成后关闭 Excel 应用程序
        if excel_app:
            try:
                excel_app.Quit()
                print("Excel 应用程序已成功关闭.")
            except Exception as ce:
                print(f"关闭 Excel 时发生错误: {ce}")
                log_error(f"关闭 Excel 时发生错误: {ce}")

        pythoncom.CoUninitialize()