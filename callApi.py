import requests
import pandas as pd
import os
import time
import io
import sys


# 设置目标 URL
url = "http://spica-tianji-dev.internal.ingka-dt.cn/maxcompute/market/hfb/performance"

# 设置请求头
headers = {
    'Content-Type': 'application/json',  # 设置请求体格式为 JSON
}

# 自定义文件路径
file_path = r'C:\RPAData\api_result.csv'

# 初始化数据列表和分页参数
all_data = []
pageNum = 1
pageSize = 100  # 设置每一页的数据量

# 主循环
while True:
    # 初始化是否有更多数据标志
    more_data = True

    # 设置请求体
    payload = {
        "pageNum": pageNum,
        "pageSize": pageSize,
        "table": "rpt_business_omni_market_hfb_performance_daily_vw",
        "instanceId": ""
    }

    while True:  # 内层循环进行请求
        try:
            # 发送 POST 请求
            response = requests.post(url, headers=headers, json=payload, timeout=10)
            response.raise_for_status()  # 检查请求是否成功

            # 解析 JSON 响应
            result = response.json()
            if result.get("success"):
                # 提取数据
                data = result.get("data", [])
                all_data.extend(data)  # 合并数据

                # 检查是否还有更多数据
                if len(data) < pageSize:  # 如果返回的数据少于 pageSize，意味着没有更多数据
                    print(f"Done A total of  {len(all_data)} data items ")
                    more_data = False  # 没有更多数据
                else:
                    pageNum += 1  # 增加页数以获取下一页
                    print(f"Receive Data - page {pageNum}")

                break  # 成功获取数据，退出此层循环

            else:
                print(f"Request Failed: {result.get('message')}")
                break  # 退出此层循环以重新尝试请求

        except Exception as err:  # 捕获所有异常
            print(f"Request Error: {err}, Retry in 5 seconds...")
            time.sleep(5)  # 重试前等待5秒

    if not more_data:  # 如果没有更多数据，退出主循环
        break

# 处理和保存所有已获取数据
if all_data:
    # 创建 DataFrame
    df = pd.DataFrame(all_data)

    # 确保目标目录存在，若不存在则创建
    os.makedirs(os.path.dirname(file_path), exist_ok=True)

    # 保存为 CSV 文件
    df.to_csv(file_path, index=False, encoding='utf-8-sig')
    print(f"Data saved {file_path}")
else:
    print("None Data")