import subprocess
import sys
import io

# 设置 stdout 编码为 utf-8
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def run_script(script_name, log_file):
    try:
        result = subprocess.run(
            [r'C:\Users\tonxu1\OneDrive - IKEA\Documents\GitHub\DailyPulseSend\.venv\Scripts\python.exe', script_name],
            capture_output=True, text=True
        )

        # 打印和记录脚本输出
        if result.stdout:
            print(result.stdout, flush=True)
            log_file.write(f"Output from {script_name}:\n{result.stdout}\n")

        # 检查是否有 stderr 输出
        if result.stderr:
            print("Error:", result.stderr, flush=True)
            # 将 stderr 输出记录到同一个日志文件
            log_file.write(f"Error in {script_name}:\n{result.stderr}\n")

        # 检查返回码
        if result.returncode != 0:
            log_file.write(f"Script {script_name} failed with return code {result.returncode}\n")

    except FileNotFoundError:
        error_message = f"File not found: {script_name}\n"
        log_file.write(error_message)
        print(error_message, flush=True)
    except Exception as e:
        error_message = f"An unexpected error occurred while running {script_name}: {str(e)}\n"
        log_file.write(error_message)
        print(error_message, flush=True)

if __name__ == "__main__":
    # 打开日志文件
    with open('log.txt', 'w', encoding='utf-8') as log_file:
        script_paths = [
            r'C:\Users\tonxu1\OneDrive - IKEA\Documents\GitHub\DailyPulseSend\callApi.py',
            r'C:\Users\tonxu1\OneDrive - IKEA\Documents\GitHub\DailyPulseSend\excelFilterSto401.py',
            r'C:\Users\tonxu1\OneDrive - IKEA\Documents\GitHub\DailyPulseSend\screenshotSendWecom.py'
        ]

        # 运行每个脚本并捕获输出和错误
        for script in script_paths:
            run_script(script, log_file)