import requests
import pandas as pd
import os

# 定义API URL
api_url = "https://api.example.com/data"  # 替换为实际API

# 定义保存文件的路径
save_path = "path/to/save/output.csv"  # 替换为你想要的保存路径

# 确保目录存在
os.makedirs(os.path.dirname(save_path), exist_ok=True)

# 调用API
response = requests.get(api_url)

# 检查响应状态码
if response.status_code == 200:
    # 解析JSON数据
    data = response.json()

    # 将数据转换为DataFrame
    df = pd.DataFrame(data)

    # 保存为CSV文件
    df.to_csv(save_path, index=False)
    print(f"数据已保存为 {save_path}")
else:
    print(f"请求失败，状态码：{response.status_code}")