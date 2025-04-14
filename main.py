# main.py
import subprocess
import sys
import io

# 设置 stdout 编码为 utf-8
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def run_script(script_name):
    result = subprocess.run([r'C:\Users\tonxu1\PycharmProjects\DailyPulseSend\.venv\Scripts\python.exe', script_name],
                            capture_output=True, text=True, encoding='utf-8')
    #print(result.stdout)  # 打印脚本输出
    if result.stderr:
        print(result.stderr)  # 打印错误信息 (如果有)


if __name__ == "__main__":
    #api
    #run_script(r'C:\Users\tonxu1\OneDrive - IKEA\Documents\GitHub\DailyPulseSend\callApi.py')

    #Excel筛选
    run_script(r'C:\Users\tonxu1\OneDrive - IKEA\Documents\GitHub\DailyPulseSend\excelFilterSto401.py')

    #截图
    run_script(r'C:\Users\tonxu1\OneDrive - IKEA\Documents\GitHub\DailyPulseSend\screenshotSendWecom.py')
