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

#以下为多张表截图不同坐标发送不同群的测试代码---------------------------------------------------------

# 设置 stdout 和 stderr 编码为 utf-8
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',errors='ignore')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8',errors='ignore')
if os.name == 'nt':
    os.system('chcp 65001')

# URL
files_and_urls = [
    (r'C:\RPAData\output-401.xlsx',
     'https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=60a349f2-f413-4f6a-a713-205a8df3fa1d',
     [(55, 306, 738, 830), (843, 306, 1643, 878)]),

    (r'C:\RPAData\output-401-2.xlsx',
     'https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=60a349f2-f413-4f6a-a713-205a8df3fa1d',
     [(25, 320, 825, 920)])
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
            raise TimeoutError("Excel is not ready within the specified time.")
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

def process_excel_file(excel_app, file_path, webhook_url, screenshot_areas):
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

        # 使用传入的截图区域
        screenshots = take_screenshots(screenshot_areas)

        for file in screenshots:
            send_image(file, webhook_url)

    except Exception as e:
        print(f"Error: {e}")
        log_error(f"Error: {e}")

    finally:
        if workbook:
            try:
                excel_app.DisplayAlerts = False  # 禁用所有警告对话框
                # 尝试关闭工作簿
                workbook.Close(SaveChanges=False)
                print('Workbook has been successfully closed.')
            except Exception as ce:
                print(f"An error occurred while closing the workbook.: {ce}")
                log_error(f"An error occurred while closing the workbook.: {ce}")

        # 删除临时文件
        for file in screenshots:
            if os.path.exists(file):
                os.remove(file)
        print("Temporary files have been deleted.")

if __name__ == "__main__":
    pythoncom.CoInitialize()
    excel_app = None

    try:
        # 启动 Excel 应用
        excel_app = win32com.client.Dispatch('Excel.Application')
        excel_app.Visible = True  # 设置 Excel 可见

        for file_path, webhook_url, screenshot_areas in files_and_urls:
            process_excel_file(excel_app, file_path, webhook_url, screenshot_areas)

    except Exception as e:
        print(f"error: {e}")
        log_error(f"error: {e}")

    finally:
        # 在所有处理完成后关闭 Excel 应用程序
        if excel_app:
            try:
                excel_app.Quit()
                print("Excel application has been successfully closed.")
            except Exception as ce:
                print(f"An error occurred while closing Excel: {ce}")
                log_error(f"An error occurred while closing Excel: {ce}")

        pythoncom.CoUninitialize()