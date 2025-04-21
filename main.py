import subprocess
import sys
import io

# 设置 stdout 编码为 utf-8
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def run_script(script_name, error_file):
    try:
        result = subprocess.run(
            [r'C:\Users\tonxu1\PycharmProjects\DailyPulseSend\.venv\Scripts\python.exe', script_name],
            capture_output=True, text=True
        )

        # 打印脚本输出
        if result.stdout:
            print(result.stdout, flush=True)

        # 检查是否有stderr输出
        if result.stderr:
            print("Error:", result.stderr, flush=True)
            # 将 stderr 输出记录到错误文件
            error_file.write(f"Error in {script_name}:\n{result.stderr}\n")

        # 检查返回码
        if result.returncode != 0:
            error_file.write(f"Script {script_name} failed with return code {result.returncode}\n")
    except Exception as e:
        error_message = f"An exception occurred while running {script_name}: {e}\n"
        print(error_message, flush=True)
        error_file.write(error_message)

if __name__ == "__main__":
    # 打开错误日志文件
    with open('error_log.txt', 'w', encoding='utf-8') as error_file:
        script_paths = [
            r'C:\Users\tonxu1\OneDrive - IKEA\Documents\GitHub\DailyPulseSend\callApi.py',
            r'C:\Users\tonxu1\OneDrive - IKEA\Documents\GitHub\DailyPulseSend\excelFilterSto401.py',
            r'C:\Users\tonxu1\OneDrive - IKEA\Documents\GitHub\DailyPulseSend\screenshotSendWecom.py'
        ]

        # 运行每个脚本并捕获错误
        for script in script_paths:
            run_script(script, error_file)