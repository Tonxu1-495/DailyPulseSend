import requests
import pandas as pd
import os
import time

# 设置目标 URL
url = "http://spica-tianji-dev.internal.ingka-dt.cn/maxcompute/market/hfb/performance"

# 设置请求头
headers = {
    'Content-Type': 'application/json',  # 设置请求体格式为 JSON
}

# 设置请求体
payload = {
    "pageNum": 1,  # 页数
    "pageSize": 10,  # 行数
    "table": "rpt_business_omni_market_hfb_performance_daily_vw",  # 表或视图
    "instanceId": ""  # 实例ID，第一次执行为空
}

# 自定义文件路径
file_path = r'C:\RPAData\hfb_performance.csv'  # 替换为您的目标路径

while True:
    try:
        # 发送 POST 请求
        response = requests.post(url, headers=headers, json=payload, timeout=10)  # 设置超时时间

        # 检查请求是否成功
        response.raise_for_status()  # 如果响应状态码不是 200，将引发 HTTPError

        result = response.json()
        if result.get("success"):
            # 提取数据
            data = result.get("data", [])

            # 创建 DataFrame
            df = pd.DataFrame(data)

            # 确保目标目录存在，若不存在则创建
            os.makedirs(os.path.dirname(file_path), exist_ok=True)

            # 保存为 CSV 文件
            df.to_csv(file_path, index=False, encoding='utf-8-sig')
            print(f"数据已成功保存为 {file_path}")
            break  # 退出循环
        else:
            print("请求失败:", result.get("message"))

    except requests.exceptions.HTTPError as http_err:
        print(f"HTTP 请求错误: {http_err}")
    except requests.exceptions.RequestException as req_err:
        print(f"请求异常: {req_err}")

    # 请求失败后的等待
    print("等待 5 分钟后重试...")
    time.sleep(300)  # 等待 5 分钟