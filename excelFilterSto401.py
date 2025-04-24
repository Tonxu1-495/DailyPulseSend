import pandas as pd
import numpy as np
from openpyxl import load_workbook
from datetime import datetime
import subprocess

# 获取当前日期和时间
now = datetime.now()
formatted_date = now.strftime("%Y-%m-%d")
#设置店号
store_code = 401


print("Current date:" + formatted_date)

#读取数据-----------------------------------------------------------------------------------------------------------------------
# 指定文件路径
file_path1 = r'C:\RPAData\api_result.csv'
file_path2 = r'C:\RPAData\output-401.xlsx'


try:
    excel1_df = pd.read_csv('C:\\RPAData\\api_result.csv', encoding='utf-8',encoding_errors='ignore')
except UnicodeDecodeError:
    excel1_df = pd.read_csv('C:\\RPAData\\api_result.csv', encoding='ISO-8859-1')  # 如果 UTF-8 失败，尝试 ISO-8859-1



# 检查 excel1_df 是否为空，并且是否包含 agg_date 列
if 'agg_date' in excel1_df.columns and not excel1_df.empty:
    # 处理 agg_date 列为日期类型
    try:
        excel1_df['agg_date'] = pd.to_datetime(excel1_df['agg_date'], format='%Y%m%d', errors='coerce')

        # 再次检查是否成功转换为日期
        if not excel1_df['agg_date'].isnull().all():
            # 获取 agg_date 列的第一行的有效日期
            first_valid_date = excel1_df['agg_date'].dropna().iloc[0]  # 获取有效日期后的第一行

            # 格式化日期
            formatted_first_agg_date = first_valid_date.strftime("%Y-%m-%d")
            print("Formatted first agg_date:", formatted_first_agg_date)
        else:
            print("No valid dates found in 'agg_date'.")

    except Exception as e:
        print("An error occurred while processing dates:", e)
else:
    print("DataFrame is empty or 'agg_date' column does not exist.")




#检索数据集合------------------------------------------------------------------------------------------------------------------------------
#offline store
filtered_df_offline_store = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (excel1_df['product_type'] == 'goods') & (excel1_df['product_type'] != 'other')].copy()
#offline store_tyd
filtered_df_offline_store_ytd = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (excel1_df['product_type'] == 'goods') & (excel1_df['product_type'] != 'other')].copy()

#online store
filtered_df_online_store = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (excel1_df['product_type'] == 'goods') & (excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()
#IKEA Food
filtered_df_ikeaFood = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['product_type'] == 'food')].copy()
#offline service
filtered_df_service_offline = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (excel1_df['hfb_no'] == 97)].copy()
#online service
filtered_df_service_online = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (excel1_df['hfb_no'] == 97)].copy()
# HFB01 offline
filtered_df_01offline = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (excel1_df['hfb_no'] == 1)].copy()
#HFB01 online
filtered_df_01online = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (excel1_df['hfb_no'] == 1) & (excel1_df['product_type'] == 'goods') & (excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()
# HFB02 offline
filtered_df_02offline = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (excel1_df['hfb_no'] == 2)].copy()
#HFB02 online
filtered_df_02online = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (excel1_df['hfb_no'] == 2) & (excel1_df['product_type'] == 'goods') & (excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()
# HFB03 offline
filtered_df_03offline = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (excel1_df['hfb_no'] == 3)].copy()
#HFB03 online
filtered_df_03online = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (excel1_df['hfb_no'] == 3) & (excel1_df['product_type'] == 'goods') & (excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()
# HFB04 offline
filtered_df_04offline = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (excel1_df['hfb_no'] == 4)].copy()
#HFB04 online
filtered_df_04online = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (excel1_df['hfb_no'] == 4) & (excel1_df['product_type'] == 'goods') & (excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()
# HFB05 offline
filtered_df_05offline = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (excel1_df['hfb_no'] == 5)].copy()
#HFB05 online
filtered_df_05online = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (excel1_df['hfb_no'] == 5) & (excel1_df['product_type'] == 'goods') & (excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()
# HFB06 offline
filtered_df_06offline = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (excel1_df['hfb_no'] == 6)].copy()
#HFB06 online
filtered_df_06online = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (excel1_df['hfb_no'] == 6) & (excel1_df['product_type'] == 'goods') & (excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()
# HFB07 offline
filtered_df_07offline = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (excel1_df['hfb_no'] == 7)].copy()
#HFB07 online
filtered_df_07online = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (excel1_df['hfb_no'] == 7) & (excel1_df['product_type'] == 'goods') & (excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()
# HFB08 offline
filtered_df_08offline = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (excel1_df['hfb_no'] == 8)].copy()
#HFB08 online
filtered_df_08online = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (excel1_df['hfb_no'] == 8) & (excel1_df['product_type'] == 'goods') & (excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()
# HFB09 offline
filtered_df_09offline = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (excel1_df['hfb_no'] == 9)].copy()
#HFB09 online
filtered_df_09online = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (excel1_df['hfb_no'] == 9) & (excel1_df['product_type'] == 'goods') & (excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()
# HFB10 offline
filtered_df_10offline = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (excel1_df['hfb_no'] == 10)].copy()
#HFB10 online
filtered_df_10online = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (excel1_df['hfb_no'] == 10) & (excel1_df['product_type'] == 'goods') & (excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()
# HFB11 offline
filtered_df_11offline = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (excel1_df['hfb_no'] == 11)].copy()
#HFB11 online
filtered_df_11online = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (excel1_df['hfb_no'] == 11) & (excel1_df['product_type'] == 'goods') & (excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()
# HFB12 offline
filtered_df_12offline = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (excel1_df['hfb_no'] == 12)].copy()
#HFB12 online
filtered_df_12online = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (excel1_df['hfb_no'] == 12) & (excel1_df['product_type'] == 'goods') & (excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()
# HFB13 offline
filtered_df_13offline = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (excel1_df['hfb_no'] == 13)].copy()
#HFB13 online
filtered_df_13online = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (excel1_df['hfb_no'] == 13) & (excel1_df['product_type'] == 'goods') & (excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()
# HFB14 offline
filtered_df_14offline = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (excel1_df['hfb_no'] == 14)].copy()
#HFB14 online
filtered_df_14online = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (excel1_df['hfb_no'] == 14) & (excel1_df['product_type'] == 'goods') & (excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()
# HFB15 offline
filtered_df_15offline = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (excel1_df['hfb_no'] == 15)].copy()
#HFB15 online
filtered_df_15online = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (excel1_df['hfb_no'] == 15) & (excel1_df['product_type'] == 'goods') & (excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()
# HFB16 offline
filtered_df_16offline = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (excel1_df['hfb_no'] == 16)].copy()
#HFB16 online
filtered_df_16online = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (excel1_df['hfb_no'] == 16) & (excel1_df['product_type'] == 'goods') & (excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()
# HFB17 offline
filtered_df_17offline = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (excel1_df['hfb_no'] == 17)].copy()
#HFB17 online
filtered_df_17online = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (excel1_df['hfb_no'] == 17) & (excel1_df['product_type'] == 'goods') & (excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()
# HFB18 offline
filtered_df_18offline = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (excel1_df['hfb_no'] == 18)].copy()
#HFB18 online
filtered_df_18online = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (excel1_df['hfb_no'] == 18) & (excel1_df['product_type'] == 'goods') & (excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()
# HFB70 offline
filtered_df_70offline = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (excel1_df['hfb_no'] == 70)].copy()
#HFB70 online
filtered_df_70online = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (excel1_df['hfb_no'] == 70) & (excel1_df['product_type'] == 'goods') & (excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()
# HFB95 offline
filtered_df_95offline = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (excel1_df['hfb_no'] == 95)].copy()
#HFB95 online
filtered_df_95online = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (excel1_df['hfb_no'] == 95) & (excel1_df['product_type'] == 'goods') & (excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()




#取值并取整-------------------------------------------------------------------------------------------------------------------------------
#HFB01offline
amt_01offline = np.round(filtered_df_01offline['sale_net_amt'].values / 1000).astype(int)
ly_amt_01offline = np.round(filtered_df_01offline['ly_sale_net_amt'].values / 1000).astype(int)
goal_01offline = np.round(filtered_df_01offline['sales_goal'].values / 1000).astype(int)
index_01_offline = np.round(filtered_df_01offline['sale_net_amt'].values) / np.round(filtered_df_01offline['sales_goal'].values) * 100
index_ly_01_offline = np.round(filtered_df_01offline['sale_net_amt'].values) / np.round(filtered_df_01offline['ly_sale_net_amt'].values) * 100
#HFB01 online
amt_01online = np.round(filtered_df_01online['sale_net_amt'].sum() / 1000).astype(int)
goal_01online = np.round(filtered_df_01online['sales_goal'].sum() / 1000).astype(int)
ly_amt_01online = np.round(filtered_df_01online['ly_sale_net_amt'].sum() / 1000).astype(int)
index_01_online = np.round(filtered_df_01online['sale_net_amt'].sum()) / np.round(filtered_df_01online['sales_goal'].sum()) * 100
index_ly_01_online = np.round(filtered_df_01online['sale_net_amt'].sum()) / np.round(filtered_df_01online['ly_sale_net_amt'].sum()) * 100
#HFB02offline
amt_02offline = np.round(filtered_df_02offline['sale_net_amt'].values / 1000).astype(int)
ly_amt_02offline = np.round(filtered_df_02offline['ly_sale_net_amt'].values / 1000).astype(int)
goal_02offline = np.round(filtered_df_02offline['sales_goal'].values / 1000).astype(int)
index_02_offline = np.round(filtered_df_02offline['sale_net_amt'].values) / np.round(filtered_df_02offline['sales_goal'].values) * 100
index_ly_02_offline = np.round(filtered_df_02offline['sale_net_amt'].values) / np.round(filtered_df_02offline['ly_sale_net_amt'].values) * 100
#HFB02 online
amt_02online = np.round(filtered_df_02online['sale_net_amt'].sum() / 1000).astype(int)
goal_02online = np.round(filtered_df_02online['sales_goal'].sum() / 1000).astype(int)
ly_amt_02online = np.round(filtered_df_02online['ly_sale_net_amt'].sum() / 1000).astype(int)
index_02_online = np.round(filtered_df_02online['sale_net_amt'].sum()) / np.round(filtered_df_02online['sales_goal'].sum()) * 100
index_ly_02_online = np.round(filtered_df_02online['sale_net_amt'].sum()) / np.round(filtered_df_02online['ly_sale_net_amt'].sum()) * 100
#HFB03offline
amt_03offline = np.round(filtered_df_03offline['sale_net_amt'].values / 1000).astype(int)
ly_amt_03offline = np.round(filtered_df_03offline['ly_sale_net_amt'].values / 1000).astype(int)
goal_03offline = np.round(filtered_df_03offline['sales_goal'].values / 1000).astype(int)
index_03_offline = np.round(filtered_df_03offline['sale_net_amt'].values) / np.round(filtered_df_03offline['sales_goal'].values) * 100
index_ly_03_offline = np.round(filtered_df_03offline['sale_net_amt'].values) / np.round(filtered_df_03offline['ly_sale_net_amt'].values) * 100
#HFB03 online
amt_03online = np.round(filtered_df_03online['sale_net_amt'].sum() / 1000).astype(int)
goal_03online = np.round(filtered_df_03online['sales_goal'].sum() / 1000).astype(int)
ly_amt_03online = np.round(filtered_df_03online['ly_sale_net_amt'].sum() / 1000).astype(int)
index_03_online = np.round(filtered_df_03online['sale_net_amt'].sum()) / np.round(filtered_df_03online['sales_goal'].sum()) * 100
index_ly_03_online = np.round(filtered_df_03online['sale_net_amt'].sum()) / np.round(filtered_df_03online['ly_sale_net_amt'].sum()) * 100
#HFB04offline
amt_04offline = np.round(filtered_df_04offline['sale_net_amt'].values / 1000).astype(int)
ly_amt_04offline = np.round(filtered_df_04offline['ly_sale_net_amt'].values / 1000).astype(int)
goal_04offline = np.round(filtered_df_04offline['sales_goal'].values / 1000).astype(int)
index_04_offline = np.round(filtered_df_04offline['sale_net_amt'].values) / np.round(filtered_df_04offline['sales_goal'].values) * 100
index_ly_04_offline = np.round(filtered_df_04offline['sale_net_amt'].values) / np.round(filtered_df_04offline['ly_sale_net_amt'].values) * 100
#HFB04 online
amt_04online = np.round(filtered_df_04online['sale_net_amt'].sum() / 1000).astype(int)
goal_04online = np.round(filtered_df_04online['sales_goal'].sum() / 1000).astype(int)
ly_amt_04online = np.round(filtered_df_04online['ly_sale_net_amt'].sum() / 1000).astype(int)
index_04_online = np.round(filtered_df_04online['sale_net_amt'].sum()) / np.round(filtered_df_04online['sales_goal'].sum()) * 100
index_ly_04_online = np.round(filtered_df_04online['sale_net_amt'].sum()) / np.round(filtered_df_04online['ly_sale_net_amt'].sum()) * 100
#HFB05offline
amt_05offline = np.round(filtered_df_05offline['sale_net_amt'].values / 1000).astype(int)
ly_amt_05offline = np.round(filtered_df_05offline['ly_sale_net_amt'].values / 1000).astype(int)
goal_05offline = np.round(filtered_df_05offline['sales_goal'].values / 1000).astype(int)
index_05_offline = np.round(filtered_df_05offline['sale_net_amt'].values) / np.round(filtered_df_05offline['sales_goal'].values) * 100
index_ly_05_offline = np.round(filtered_df_05offline['sale_net_amt'].values) / np.round(filtered_df_05offline['ly_sale_net_amt'].values) * 100
#HFB05 online
amt_05online = np.round(filtered_df_05online['sale_net_amt'].sum() / 1000).astype(int)
goal_05online = np.round(filtered_df_05online['sales_goal'].sum() / 1000).astype(int)
ly_amt_05online = np.round(filtered_df_05online['ly_sale_net_amt'].sum() / 1000).astype(int)
index_05_online = np.round(filtered_df_05online['sale_net_amt'].sum()) / np.round(filtered_df_05online['sales_goal'].sum()) * 100
index_ly_05_online = np.round(filtered_df_05online['sale_net_amt'].sum()) / np.round(filtered_df_05online['ly_sale_net_amt'].sum()) * 100
#HFB06offline
amt_06offline = np.round(filtered_df_06offline['sale_net_amt'].values / 1000).astype(int)
ly_amt_06offline = np.round(filtered_df_06offline['ly_sale_net_amt'].values / 1000).astype(int)
goal_06offline = np.round(filtered_df_06offline['sales_goal'].values / 1000).astype(int)
index_06_offline = np.round(filtered_df_06offline['sale_net_amt'].values) / np.round(filtered_df_06offline['sales_goal'].values) * 100
index_ly_06_offline = np.round(filtered_df_06offline['sale_net_amt'].values) / np.round(filtered_df_06offline['ly_sale_net_amt'].values) * 100
#HFB06 online
amt_06online = np.round(filtered_df_06online['sale_net_amt'].sum() / 1000).astype(int)
goal_06online = np.round(filtered_df_06online['sales_goal'].sum() / 1000).astype(int)
ly_amt_06online = np.round(filtered_df_06online['ly_sale_net_amt'].sum() / 1000).astype(int)
index_06_online = np.round(filtered_df_06online['sale_net_amt'].sum()) / np.round(filtered_df_06online['sales_goal'].sum()) * 100
index_ly_06_online = np.round(filtered_df_06online['sale_net_amt'].sum()) / np.round(filtered_df_06online['ly_sale_net_amt'].sum()) * 100
# HFB07 offline
amt_07offline = np.round(filtered_df_07offline['sale_net_amt'].values / 1000).astype(int)
ly_amt_07offline = np.round(filtered_df_07offline['ly_sale_net_amt'].values / 1000).astype(int)
goal_07offline = np.round(filtered_df_07offline['sales_goal'].values / 1000).astype(int)
index_07_offline = np.round(filtered_df_07offline['sale_net_amt'].values) / np.round(filtered_df_07offline['sales_goal'].values) * 100
index_ly_07_offline = np.round(filtered_df_07offline['sale_net_amt'].values) / np.round(filtered_df_07offline['ly_sale_net_amt'].values) * 100

# HFB07 online
amt_07online = np.round(filtered_df_07online['sale_net_amt'].sum() / 1000).astype(int)
goal_07online = np.round(filtered_df_07online['sales_goal'].sum() / 1000).astype(int)
ly_amt_07online = np.round(filtered_df_07online['ly_sale_net_amt'].sum() / 1000).astype(int)
index_07_online = np.round(filtered_df_07online['sale_net_amt'].sum()) / np.round(filtered_df_07online['sales_goal'].sum()) * 100
index_ly_07_online = np.round(filtered_df_07online['sale_net_amt'].sum()) / np.round(filtered_df_07online['ly_sale_net_amt'].sum()) * 100

# HFB08 offline
amt_08offline = np.round(filtered_df_08offline['sale_net_amt'].values / 1000).astype(int)
ly_amt_08offline = np.round(filtered_df_08offline['ly_sale_net_amt'].values / 1000).astype(int)
goal_08offline = np.round(filtered_df_08offline['sales_goal'].values / 1000).astype(int)
index_08_offline = np.round(filtered_df_08offline['sale_net_amt'].values) / np.round(filtered_df_08offline['sales_goal'].values) * 100
index_ly_08_offline = np.round(filtered_df_08offline['sale_net_amt'].values) / np.round(filtered_df_08offline['ly_sale_net_amt'].values) * 100

# HFB08 online
amt_08online = np.round(filtered_df_08online['sale_net_amt'].sum() / 1000).astype(int)
goal_08online = np.round(filtered_df_08online['sales_goal'].sum() / 1000).astype(int)
ly_amt_08online = np.round(filtered_df_08online['ly_sale_net_amt'].sum() / 1000).astype(int)
index_08_online = np.round(filtered_df_08online['sale_net_amt'].sum()) / np.round(filtered_df_08online['sales_goal'].sum()) * 100
index_ly_08_online = np.round(filtered_df_08online['sale_net_amt'].sum()) / np.round(filtered_df_08online['ly_sale_net_amt'].sum()) * 100

# HFB09 offline
amt_09offline = np.round(filtered_df_09offline['sale_net_amt'].values / 1000).astype(int)
ly_amt_09offline = np.round(filtered_df_09offline['ly_sale_net_amt'].values / 1000).astype(int)
goal_09offline = np.round(filtered_df_09offline['sales_goal'].values / 1000).astype(int)
index_09_offline = np.round(filtered_df_09offline['sale_net_amt'].values) / np.round(filtered_df_09offline['sales_goal'].values) * 100
index_ly_09_offline = np.round(filtered_df_09offline['sale_net_amt'].values) / np.round(filtered_df_09offline['ly_sale_net_amt'].values) * 100

# HFB09 online
amt_09online = np.round(filtered_df_09online['sale_net_amt'].sum() / 1000).astype(int)
goal_09online = np.round(filtered_df_09online['sales_goal'].sum() / 1000).astype(int)
ly_amt_09online = np.round(filtered_df_09online['ly_sale_net_amt'].sum() / 1000).astype(int)
index_09_online = np.round(filtered_df_09online['sale_net_amt'].sum()) / np.round(filtered_df_09online['sales_goal'].sum()) * 100
index_ly_09_online = np.round(filtered_df_09online['sale_net_amt'].sum()) / np.round(filtered_df_09online['ly_sale_net_amt'].sum()) * 100

# HFB10 offline
amt_10offline = np.round(filtered_df_10offline['sale_net_amt'].values / 1000).astype(int)
ly_amt_10offline = np.round(filtered_df_10offline['ly_sale_net_amt'].values / 1000).astype(int)
goal_10offline = np.round(filtered_df_10offline['sales_goal'].values / 1000).astype(int)
index_10_offline = np.round(filtered_df_10offline['sale_net_amt'].values) / np.round(filtered_df_10offline['sales_goal'].values) * 100
index_ly_10_offline = np.round(filtered_df_10offline['sale_net_amt'].values) / np.round(filtered_df_10offline['ly_sale_net_amt'].values) * 100

# HFB10 online
amt_10online = np.round(filtered_df_10online['sale_net_amt'].sum() / 1000).astype(int)
goal_10online = np.round(filtered_df_10online['sales_goal'].sum() / 1000).astype(int)
ly_amt_10online = np.round(filtered_df_10online['ly_sale_net_amt'].sum() / 1000).astype(int)
index_10_online = np.round(filtered_df_10online['sale_net_amt'].sum()) / np.round(filtered_df_10online['sales_goal'].sum()) * 100
index_ly_10_online = np.round(filtered_df_10online['sale_net_amt'].sum()) / np.round(filtered_df_10online['ly_sale_net_amt'].sum()) * 100

# HFB11 offline
amt_11offline = np.round(filtered_df_11offline['sale_net_amt'].values / 1000).astype(int)
ly_amt_11offline = np.round(filtered_df_11offline['ly_sale_net_amt'].values / 1000).astype(int)
goal_11offline = np.round(filtered_df_11offline['sales_goal'].values / 1000).astype(int)
index_11_offline = np.round(filtered_df_11offline['sale_net_amt'].values) / np.round(filtered_df_11offline['sales_goal'].values) * 100
index_ly_11_offline = np.round(filtered_df_11offline['sale_net_amt'].values) / np.round(filtered_df_11offline['ly_sale_net_amt'].values) * 100

# HFB11 online
amt_11online = np.round(filtered_df_11online['sale_net_amt'].sum() / 1000).astype(int)
goal_11online = np.round(filtered_df_11online['sales_goal'].sum() / 1000).astype(int)
ly_amt_11online = np.round(filtered_df_11online['ly_sale_net_amt'].sum() / 1000).astype(int)
index_11_online = np.round(filtered_df_11online['sale_net_amt'].sum()) / np.round(filtered_df_11online['sales_goal'].sum()) * 100
index_ly_11_online = np.round(filtered_df_11online['sale_net_amt'].sum()) / np.round(filtered_df_11online['ly_sale_net_amt'].sum()) * 100

# HFB12 offline
amt_12offline = np.round(filtered_df_12offline['sale_net_amt'].values / 1000).astype(int)
ly_amt_12offline = np.round(filtered_df_12offline['ly_sale_net_amt'].values / 1000).astype(int)
goal_12offline = np.round(filtered_df_12offline['sales_goal'].values / 1000).astype(int)
index_12_offline = np.round(filtered_df_12offline['sale_net_amt'].values) / np.round(filtered_df_12offline['sales_goal'].values) * 100
index_ly_12_offline = np.round(filtered_df_12offline['sale_net_amt'].values) / np.round(filtered_df_12offline['ly_sale_net_amt'].values) * 100

# HFB12 online
amt_12online = np.round(filtered_df_12online['sale_net_amt'].sum() / 1000).astype(int)
goal_12online = np.round(filtered_df_12online['sales_goal'].sum() / 1000).astype(int)
ly_amt_12online = np.round(filtered_df_12online['ly_sale_net_amt'].sum() / 1000).astype(int)
index_12_online = np.round(filtered_df_12online['sale_net_amt'].sum()) / np.round(filtered_df_12online['sales_goal'].sum()) * 100
index_ly_12_online = np.round(filtered_df_12online['sale_net_amt'].sum()) / np.round(filtered_df_12online['ly_sale_net_amt'].sum()) * 100

# HFB13 offline
amt_13offline = np.round(filtered_df_13offline['sale_net_amt'].values / 1000).astype(int)
ly_amt_13offline = np.round(filtered_df_13offline['ly_sale_net_amt'].values / 1000).astype(int)
goal_13offline = np.round(filtered_df_13offline['sales_goal'].values / 1000).astype(int)
index_13_offline = np.round(filtered_df_13offline['sale_net_amt'].values) / np.round(filtered_df_13offline['sales_goal'].values) * 100
index_ly_13_offline = np.round(filtered_df_13offline['sale_net_amt'].values) / np.round(filtered_df_13offline['ly_sale_net_amt'].values) * 100

# HFB13 online
amt_13online = np.round(filtered_df_13online['sale_net_amt'].sum() / 1000).astype(int)
goal_13online = np.round(filtered_df_13online['sales_goal'].sum() / 1000).astype(int)
ly_amt_13online = np.round(filtered_df_13online['ly_sale_net_amt'].sum() / 1000).astype(int)
index_13_online = np.round(filtered_df_13online['sale_net_amt'].sum()) / np.round(filtered_df_13online['sales_goal'].sum()) * 100
index_ly_13_online = np.round(filtered_df_13online['sale_net_amt'].sum()) / np.round(filtered_df_13online['ly_sale_net_amt'].sum()) * 100

# HFB14 offline
amt_14offline = np.round(filtered_df_14offline['sale_net_amt'].values / 1000).astype(int)
ly_amt_14offline = np.round(filtered_df_14offline['ly_sale_net_amt'].values / 1000).astype(int)
goal_14offline = np.round(filtered_df_14offline['sales_goal'].values / 1000).astype(int)
index_14_offline = np.round(filtered_df_14offline['sale_net_amt'].values) / np.round(filtered_df_14offline['sales_goal'].values) * 100
index_ly_14_offline = np.round(filtered_df_14offline['sale_net_amt'].values) / np.round(filtered_df_14offline['ly_sale_net_amt'].values) * 100

# HFB14 online
amt_14online = np.round(filtered_df_14online['sale_net_amt'].sum() / 1000).astype(int)
goal_14online = np.round(filtered_df_14online['sales_goal'].sum() / 1000).astype(int)
ly_amt_14online = np.round(filtered_df_14online['ly_sale_net_amt'].sum() / 1000).astype(int)
index_14_online = np.round(filtered_df_14online['sale_net_amt'].sum()) / np.round(filtered_df_14online['sales_goal'].sum()) * 100
index_ly_14_online = np.round(filtered_df_14online['sale_net_amt'].sum()) / np.round(filtered_df_14online['ly_sale_net_amt'].sum()) * 100

# HFB15 offline
amt_15offline = np.round(filtered_df_15offline['sale_net_amt'].values / 1000).astype(int)
ly_amt_15offline = np.round(filtered_df_15offline['ly_sale_net_amt'].values / 1000).astype(int)
goal_15offline = np.round(filtered_df_15offline['sales_goal'].values / 1000).astype(int)
index_15_offline = np.round(filtered_df_15offline['sale_net_amt'].values) / np.round(filtered_df_15offline['sales_goal'].values) * 100
index_ly_15_offline = np.round(filtered_df_15offline['sale_net_amt'].values) / np.round(filtered_df_15offline['ly_sale_net_amt'].values) * 100

# HFB15 online
amt_15online = np.round(filtered_df_15online['sale_net_amt'].sum() / 1000).astype(int)
goal_15online = np.round(filtered_df_15online['sales_goal'].sum() / 1000).astype(int)
ly_amt_15online = np.round(filtered_df_15online['ly_sale_net_amt'].sum() / 1000).astype(int)
index_15_online = np.round(filtered_df_15online['sale_net_amt'].sum()) / np.round(filtered_df_15online['sales_goal'].sum()) * 100
index_ly_15_online = np.round(filtered_df_15online['sale_net_amt'].sum()) / np.round(filtered_df_15online['ly_sale_net_amt'].sum()) * 100

# HFB16 offline
amt_16offline = np.round(filtered_df_16offline['sale_net_amt'].values / 1000).astype(int)
ly_amt_16offline = np.round(filtered_df_16offline['ly_sale_net_amt'].values / 1000).astype(int)
goal_16offline = np.round(filtered_df_16offline['sales_goal'].values / 1000).astype(int)
index_16_offline = np.round(filtered_df_16offline['sale_net_amt'].values) / np.round(filtered_df_16offline['sales_goal'].values) * 100
index_ly_16_offline = np.round(filtered_df_16offline['sale_net_amt'].values) / np.round(filtered_df_16offline['ly_sale_net_amt'].values) * 100

# HFB16 online
amt_16online = np.round(filtered_df_16online['sale_net_amt'].sum() / 1000).astype(int)
goal_16online = np.round(filtered_df_16online['sales_goal'].sum() / 1000).astype(int)
ly_amt_16online = np.round(filtered_df_16online['ly_sale_net_amt'].sum() / 1000).astype(int)
index_16_online = np.round(filtered_df_16online['sale_net_amt'].sum()) / np.round(filtered_df_16online['sales_goal'].sum()) * 100
index_ly_16_online = np.round(filtered_df_16online['sale_net_amt'].sum()) / np.round(filtered_df_16online['ly_sale_net_amt'].sum()) * 100

# HFB17 offline
amt_17offline = np.round(filtered_df_17offline['sale_net_amt'].values / 1000).astype(int)
ly_amt_17offline = np.round(filtered_df_17offline['ly_sale_net_amt'].values / 1000).astype(int)
goal_17offline = np.round(filtered_df_17offline['sales_goal'].values / 1000).astype(int)
index_17_offline = np.round(filtered_df_17offline['sale_net_amt'].values) / np.round(filtered_df_17offline['sales_goal'].values) * 100
index_ly_17_offline = np.round(filtered_df_17offline['sale_net_amt'].values) / np.round(filtered_df_17offline['ly_sale_net_amt'].values) * 100

# HFB17 online
amt_17online = np.round(filtered_df_17online['sale_net_amt'].sum() / 1000).astype(int)
goal_17online = np.round(filtered_df_17online['sales_goal'].sum() / 1000).astype(int)
ly_amt_17online = np.round(filtered_df_17online['ly_sale_net_amt'].sum() / 1000).astype(int)
index_17_online = np.round(filtered_df_17online['sale_net_amt'].sum()) / np.round(filtered_df_17online['sales_goal'].sum()) * 100
index_ly_17_online = np.round(filtered_df_17online['sale_net_amt'].sum()) / np.round(filtered_df_17online['ly_sale_net_amt'].sum()) * 100

# HFB18 offline
amt_18offline = np.round(filtered_df_18offline['sale_net_amt'].values / 1000).astype(int)
ly_amt_18offline = np.round(filtered_df_18offline['ly_sale_net_amt'].values / 1000).astype(int)
goal_18offline = np.round(filtered_df_18offline['sales_goal'].values / 1000).astype(int)
index_18_offline = np.round(filtered_df_18offline['sale_net_amt'].values) / np.round(filtered_df_18offline['sales_goal'].values) * 100
index_ly_18_offline = np.round(filtered_df_18offline['sale_net_amt'].values) / np.round(filtered_df_18offline['ly_sale_net_amt'].values) * 100

# HFB18 online
amt_18online = np.round(filtered_df_18online['sale_net_amt'].sum() / 1000).astype(int)
goal_18online = np.round(filtered_df_18online['sales_goal'].sum() / 1000).astype(int)
ly_amt_18online = np.round(filtered_df_18online['ly_sale_net_amt'].sum() / 1000).astype(int)
index_18_online = np.round(filtered_df_18online['sale_net_amt'].sum()) / np.round(filtered_df_18online['sales_goal'].sum()) * 100
index_ly_18_online = np.round(filtered_df_18online['sale_net_amt'].sum()) / np.round(filtered_df_18online['ly_sale_net_amt'].sum()) * 100
#HFB70offline
amt_70offline = np.round(filtered_df_70offline['sale_net_amt'].values)
ly_amt_70offline = np.round(filtered_df_70offline['ly_sale_net_amt'].values)
goal_70offline = np.round(filtered_df_70offline['sales_goal'].values)

# 计算 index_70_offline
if goal_70offline[0] != 0:
    index_70_offline = amt_70offline[0] / goal_70offline[0] * 100
else:
    index_70_offline = 0

# 计算 index_70_offline_ly
if ly_amt_70offline[0] != 0:
    index_70_offline_ly = amt_70offline[0] / ly_amt_70offline[0] * 100
else:
    index_70_offline_ly = 0


#HFB70 online
amt_70online = np.round(filtered_df_70online['sale_net_amt'].sum())
goal_70online = np.round(filtered_df_70online['sales_goal'].sum())
ly_amt_70online = np.round(filtered_df_70online['ly_sale_net_amt'].sum())
# 计算 index_70_offline
if goal_70online != 0:
    index_70_online = amt_70online / goal_70online * 100
else:
    index_70_online = 0

# 计算 index_70_offline_ly
if ly_amt_70online != 0:
    index_70_online_ly = amt_70online / ly_amt_70online * 100
else:
    index_70_online_ly = 0

amt_70_total = amt_70offline[0] + amt_70online
if goal_70offline != 0:
    index_70_total = amt_70_total / goal_70offline * 100
else:
    index_70_total = 0



#HFB95offline
amt_95offline = np.round(filtered_df_95offline['sale_net_amt'].values / 1000).astype(int)
ly_amt_95offline = np.round(filtered_df_95offline['ly_sale_net_amt'].values / 1000).astype(int)
goal_95offline = np.round(filtered_df_95offline['sales_goal'].values / 1000).astype(int)
#HFB95 online
amt_95online = np.round(filtered_df_95online['sale_net_amt'].sum() / 1000).astype(int)
goal_95online = np.round(filtered_df_95online['sales_goal'].sum() / 1000).astype(int)
ly_amt_95online = np.round(filtered_df_95online['ly_sale_net_amt'].sum() / 1000).astype(int)

#Store_Daily
store_offline_ly = np.round(filtered_df_offline_store['ly_sale_net_amt'].sum() / 1000).astype(int)
store_offline_goal = np.round(filtered_df_offline_store['sales_goal'].sum() / 1000).astype(int)
store_offline_amt = np.round(filtered_df_offline_store['sale_net_amt'].sum() / 1000).astype(int)
store_online_ly = np.round(filtered_df_online_store['ly_sale_net_amt'].sum() / 1000).astype(int)
store_online_goal = np.round(filtered_df_online_store['sales_goal'].sum() / 1000).astype(int)
store_online_amt = np.round(filtered_df_online_store['sale_net_amt'].sum() / 1000).astype(int)
store_ikeaFood_amt = np.round(filtered_df_ikeaFood['sale_net_amt'].sum() / 1000).astype(int)
store_ikeaFood_goal = np.round(filtered_df_ikeaFood['sales_goal'].sum() / 1000).astype(int)
store_ikeaFood_ly = np.round(filtered_df_ikeaFood['ly_sale_net_amt'].sum() / 1000).astype(int)
store_offline_service_amt = np.round(filtered_df_service_offline['sale_net_amt'].sum() / 1000).astype(int)
store_offline_service_goal = np.round(filtered_df_service_offline['sales_goal'].sum() / 1000).astype(int)
store_offline_service_ly = np.round(filtered_df_service_offline['ly_sale_net_amt'].sum() / 1000).astype(int)
store_online_service_amt = np.round(filtered_df_service_online['sale_net_amt'].sum() / 1000).astype(int)
store_online_service_goal = np.round(filtered_df_service_online['sales_goal'].sum() / 1000).astype(int)
store_online_service_ly = np.round(filtered_df_service_online['ly_sale_net_amt'].sum() / 1000).astype(int)

#Store_Index
index_store_offline = np.round(filtered_df_offline_store['sale_net_amt'].sum()) / np.round(filtered_df_offline_store['sales_goal'].sum()) * 100
index_store_offline_ly = np.round(filtered_df_offline_store['sale_net_amt'].sum()) / np.round(filtered_df_offline_store['ly_sale_net_amt'].sum()) * 100
index_store_online = np.round(filtered_df_online_store['sale_net_amt'].sum()) / np.round(filtered_df_online_store['sales_goal'].sum()) * 100
index_store_online_ly = np.round(filtered_df_online_store['sale_net_amt'].sum()) / np.round(filtered_df_online_store['ly_sale_net_amt'].sum()) * 100
index_goods_total = (np.round(filtered_df_offline_store['sale_net_amt'].sum()) + np.round(filtered_df_online_store['sale_net_amt'].sum())) / (np.round(filtered_df_offline_store['sales_goal'].sum() + np.round(filtered_df_online_store['sales_goal'].sum()))) * 100
index_goods_total_ly = (np.round(filtered_df_offline_store['sale_net_amt'].sum()) + np.round(filtered_df_online_store['sale_net_amt'].sum())) / (np.round(filtered_df_offline_store['ly_sale_net_amt'].sum() + np.round(filtered_df_online_store['ly_sale_net_amt'].sum()))) * 100
index_store_service_offline = np.round(filtered_df_service_offline['sale_net_amt'].sum()) / np.round(filtered_df_service_offline['sales_goal'].sum()) * 100
index_store_service_offline_ly = np.round(filtered_df_service_offline['sale_net_amt'].sum()) / np.round(filtered_df_service_offline['ly_sale_net_amt'].sum()) * 100
index_store_service_online = np.round(filtered_df_service_online['sale_net_amt'].sum()) / np.round(filtered_df_service_online['sales_goal'].sum()) * 100
index_store_service_online_ly = np.round(filtered_df_service_online['sale_net_amt'].sum()) / np.round(filtered_df_service_online['ly_sale_net_amt'].sum()) * 100
index_service_total = (np.round(filtered_df_service_offline['sale_net_amt'].sum()) + np.round(filtered_df_service_online['sale_net_amt'].sum())) / (np.round(filtered_df_service_offline['sales_goal'].sum() + np.round(filtered_df_service_online['sales_goal'].sum()))) * 100
index_service_total_ly = (np.round(filtered_df_service_offline['sale_net_amt'].sum()) + np.round(filtered_df_service_online['sale_net_amt'].sum())) / (np.round(filtered_df_service_offline['ly_sale_net_amt'].sum() + np.round(filtered_df_service_online['ly_sale_net_amt'].sum()))) * 100

#Store_index_ytd
index_store_offline_ytd = np.round(filtered_df_offline_store_ytd['sale_net_amt_fytd'].sum()) / np.round(filtered_df_offline_store_ytd['sales_goal_fytd'].sum()) * 100
index_store_offline_ly_ytd = np.round(filtered_df_offline_store_ytd['sale_net_amt_fytd'].sum()) / np.round(filtered_df_offline_store_ytd['ly_sale_net_amt_fytd'].sum()) * 100
index_store_online_ytd = np.round(filtered_df_online_store['sale_net_amt_fytd'].sum()) / np.round(filtered_df_online_store['sales_goal_fytd'].sum()) * 100
index_store_online_ly_ytd = np.round(filtered_df_online_store['sale_net_amt_fytd'].sum()) / np.round(filtered_df_online_store['ly_sale_net_amt_fytd'].sum()) * 100
index_goods_total_ytd = (np.round(filtered_df_offline_store_ytd['sale_net_amt_fytd'].sum()) + np.round(filtered_df_online_store['sale_net_amt_fytd'].sum())) / (np.round(filtered_df_offline_store_ytd['sales_goal_fytd'].sum() + np.round(filtered_df_online_store['sales_goal_fytd'].sum()))) * 100
index_goods_total_ly_ytd = (np.round(filtered_df_offline_store_ytd['sale_net_amt_fytd'].sum()) + np.round(filtered_df_online_store['sale_net_amt_fytd'].sum())) / (np.round(filtered_df_offline_store_ytd['ly_sale_net_amt_fytd'].sum() + np.round(filtered_df_online_store['ly_sale_net_amt_fytd'].sum()))) * 100
index_store_service_offline_ytd = np.round(filtered_df_service_offline['sale_net_amt_fytd'].sum()) / np.round(filtered_df_service_offline['sales_goal_fytd'].sum()) * 100
index_store_service_offline_ly_ytd = np.round(filtered_df_service_offline['sale_net_amt_fytd'].sum()) / np.round(filtered_df_service_offline['ly_sale_net_amt_fytd'].sum()) * 100
index_store_service_online_ytd = np.round(filtered_df_service_online['sale_net_amt_fytd'].sum()) / np.round(filtered_df_service_online['sales_goal_fytd'].sum()) * 100
index_store_service_online_ly_ytd = np.round(filtered_df_service_online['sale_net_amt_fytd'].sum()) / np.round(filtered_df_service_online['ly_sale_net_amt_fytd'].sum()) * 100
index_service_total_ytd = (np.round(filtered_df_service_offline['sale_net_amt_fytd'].sum()) + np.round(filtered_df_service_online['sale_net_amt_fytd'].sum())) / (np.round(filtered_df_service_offline['sales_goal_fytd'].sum() + np.round(filtered_df_service_online['sales_goal_fytd'].sum()))) * 100
index_service_total_ly_ytd = (np.round(filtered_df_service_offline['sale_net_amt_fytd'].sum()) + np.round(filtered_df_service_online['sale_net_amt_fytd'].sum())) / (np.round(filtered_df_service_offline['ly_sale_net_amt_fytd'].sum() + np.round(filtered_df_service_online['ly_sale_net_amt_fytd'].sum()))) * 100



#Store_YTD
store_offline_ly_ytd = np.round(filtered_df_offline_store_ytd['ly_sale_net_amt_fytd'].sum() / 1000).astype(int)
store_offline_goal_ytd = np.round(filtered_df_offline_store_ytd['sales_goal_fytd'].sum() / 1000).astype(int)
store_offline_amt_ytd = np.round(filtered_df_offline_store_ytd['sale_net_amt_fytd'].sum() / 1000).astype(int)
store_online_ly_ytd = np.round(filtered_df_online_store['ly_sale_net_amt_fytd'].sum() / 1000).astype(int)
store_online_goal_ytd = np.round(filtered_df_online_store['sales_goal_fytd'].sum() / 1000).astype(int)
store_online_amt_ytd = np.round(filtered_df_online_store['sale_net_amt_fytd'].sum() / 1000).astype(int)
store_ikeaFood_amt_ytd = np.round(filtered_df_ikeaFood['sale_net_amt_fytd'].sum() / 1000).astype(int)
store_ikeaFood_goal_ytd = np.round(filtered_df_ikeaFood['sales_goal_fytd'].sum() / 1000).astype(int)
store_ikeaFood_ly_ytd = np.round(filtered_df_ikeaFood['ly_sale_net_amt_fytd'].sum() / 1000).astype(int)
store_offline_service_amt_ytd = np.round(filtered_df_service_offline['sale_net_amt_fytd'].sum() / 1000).astype(int)
store_offline_service_goal_ytd = np.round(filtered_df_service_offline['sales_goal_fytd'].sum() / 1000).astype(int)
store_offline_service_ly_ytd = np.round(filtered_df_service_offline['ly_sale_net_amt_fytd'].sum() / 1000).astype(int)
store_online_service_amt_ytd = np.round(filtered_df_service_online['sale_net_amt_fytd'].sum() / 1000).astype(int)
store_online_service_goal_ytd = np.round(filtered_df_service_online['sales_goal_fytd'].sum() / 1000).astype(int)
store_online_service_ly_ytd = np.round(filtered_df_service_online['ly_sale_net_amt_fytd'].sum() / 1000).astype(int)

# 写入数据----------------------------------------------------------------------------------------------------------
# 加载现有的 Excel 文件
workbook = load_workbook(file_path2)
# 选择要写入的工作表
sheet = workbook['HZ 日结模版']
excel_df = pd.read_excel(file_path2, sheet_name='HZ 日结模版')

#写入单元格
if len(ly_amt_01offline) > 0 and len(goal_01offline) > 0:

    #HFB Index
    sheet['P26'] = index_store_offline
    sheet['Q26'] = index_store_offline_ly
    sheet['U26'] = index_store_online
    sheet['V26'] = index_store_online_ly
    sheet['Z26'] = index_goods_total
    sheet['AA26'] = index_goods_total_ly
    # HFB01 offline
    sheet['P6'] = index_01_offline[0]
    sheet['Q6'] = index_ly_01_offline[0]
    sheet['U6'] = index_01_online
    sheet['V6'] = index_ly_01_online

    # HFB02 offline
    sheet['P7'] = index_02_offline[0]
    sheet['Q7'] = index_ly_02_offline[0]
    sheet['U7'] = index_02_online
    sheet['V7'] = index_ly_02_online

    # HFB03 offline
    sheet['P8'] = index_03_offline[0]
    sheet['Q8'] = index_ly_03_offline[0]
    sheet['U8'] = index_03_online
    sheet['V8'] = index_ly_03_online

    # HFB04 offline
    sheet['P9'] = index_04_offline[0]
    sheet['Q9'] = index_ly_04_offline[0]
    sheet['U9'] = index_04_online
    sheet['V9'] = index_ly_04_online

    # HFB05 offline
    sheet['P10'] = index_05_offline[0]
    sheet['Q10'] = index_ly_05_offline[0]
    sheet['U10'] = index_05_online
    sheet['V10'] = index_ly_05_online

    # HFB06 offline
    sheet['P11'] = index_06_offline[0]
    sheet['Q11'] = index_ly_06_offline[0]
    sheet['U11'] = index_06_online
    sheet['V11'] = index_ly_06_online

    # HFB07 offline
    sheet['P12'] = index_07_offline[0]
    sheet['Q12'] = index_ly_07_offline[0]
    sheet['U12'] = index_07_online
    sheet['V12'] = index_ly_07_online

    # HFB08 offline
    sheet['P13'] = index_08_offline[0]
    sheet['Q13'] = index_ly_08_offline[0]
    sheet['U13'] = index_08_online
    sheet['V13'] = index_ly_08_online

    # HFB09 offline
    sheet['P14'] = index_09_offline[0]
    sheet['Q14'] = index_ly_09_offline[0]
    sheet['U14'] = index_09_online
    sheet['V14'] = index_ly_09_online

    # HFB10 offline
    sheet['P15'] = index_10_offline[0]
    sheet['Q15'] = index_ly_10_offline[0]
    sheet['U15'] = index_10_online
    sheet['V15'] = index_ly_10_online

    # HFB11 offline
    sheet['P16'] = index_11_offline[0]
    sheet['Q16'] = index_ly_11_offline[0]
    sheet['U16'] = index_11_online
    sheet['V16'] = index_ly_11_online

    # HFB12 offline
    sheet['P17'] = index_12_offline[0]
    sheet['Q17'] = index_ly_12_offline[0]
    sheet['U17'] = index_12_online
    sheet['V17'] = index_ly_12_online

    # HFB13 offline
    sheet['P18'] = index_13_offline[0]
    sheet['Q18'] = index_ly_13_offline[0]
    sheet['U18'] = index_13_online
    sheet['V18'] = index_ly_13_online

    # HFB14 offline
    sheet['P19'] = index_14_offline[0]
    sheet['Q19'] = index_ly_14_offline[0]
    sheet['U19'] = index_14_online
    sheet['V19'] = index_ly_14_online

    # HFB15 offline
    sheet['P20'] = index_15_offline[0]
    sheet['Q20'] = index_ly_15_offline[0]
    sheet['U20'] = index_15_online
    sheet['V20'] = index_ly_15_online

    # HFB16 offline
    sheet['P21'] = index_16_offline[0]
    sheet['Q21'] = index_ly_16_offline[0]
    sheet['U21'] = index_16_online
    sheet['V21'] = index_ly_16_online

    # HFB17 offline
    sheet['P22'] = index_17_offline[0]
    sheet['Q22'] = index_ly_17_offline[0]
    sheet['U22'] = index_17_online
    sheet['V22'] = index_ly_17_online

    # HFB18 offline
    sheet['P23'] = index_18_offline[0]
    sheet['Q23'] = index_ly_18_offline[0]
    sheet['U23'] = index_18_online
    sheet['V23'] = index_ly_18_online

    #Store Index
    sheet['F6'] = index_store_offline
    sheet['H6'] = index_store_offline_ly
    sheet['F7'] = index_store_online
    sheet['H7'] = index_store_online_ly
    sheet['F8'] = index_goods_total
    sheet['H8'] = index_goods_total_ly
    sheet['F10'] = index_store_service_offline
    sheet['H10'] = index_store_service_offline_ly
    sheet['F11'] = index_store_service_online
    sheet['H11'] = index_store_service_online_ly
    sheet['F12'] = index_service_total
    sheet['H12'] = index_service_total_ly

    #Store_Index_YTD
    sheet['F16'] = index_store_offline_ytd
    sheet['H16'] = index_store_offline_ly_ytd
    sheet['F17'] = index_store_online_ytd
    sheet['H17'] = index_store_online_ly_ytd
    sheet['F18'] = index_goods_total_ytd
    sheet['H18'] = index_goods_total_ly_ytd
    sheet['F20'] = index_store_service_offline_ytd
    sheet['H20'] = index_store_service_offline_ly_ytd
    sheet['F21'] = index_store_service_online_ytd
    sheet['H21'] = index_store_service_online_ly_ytd
    sheet['F22'] = index_service_total_ytd
    sheet['H22'] = index_service_total_ly_ytd

    sheet['C4'] = formatted_first_agg_date # 写入当天日期
    # HFB01
    sheet['O6'] = amt_01offline[0]  # offline_amt
    sheet['N6'] = goal_01offline[0]  # offline_goal
    sheet['M6'] = ly_amt_01offline[0]  # offline_amt_ly
    sheet['R6'] = ly_amt_01online # online_amt_ly
    sheet['S6'] = goal_01online # online_goal
    sheet['T6'] = amt_01online  # online_amt
    #HFB02
    sheet['O7'] = amt_02offline[0]  # offline_amt
    sheet['N7'] = goal_02offline[0]  # offline_goal
    sheet['M7'] = ly_amt_02offline[0]  # offline_amt_ly
    sheet['R7'] = ly_amt_02online # online_amt_ly
    sheet['S7'] = goal_02online # online_goal
    sheet['T7'] = amt_02online  # online_amt
    #HFB03
    sheet['O8'] = amt_03offline[0]  # offline_amt
    sheet['N8'] = goal_03offline[0]  # offline_goal
    sheet['M8'] = ly_amt_03offline[0]  # offline_amt_ly
    sheet['R8'] = ly_amt_03online # online_amt_ly
    sheet['S8'] = goal_03online # online_goal
    sheet['T8'] = amt_03online  # online_amt
    #HFB04
    sheet['O9'] = amt_04offline[0]  # offline_amt
    sheet['N9'] = goal_04offline[0]  # offline_goal
    sheet['M9'] = ly_amt_04offline[0]  # offline_amt_ly
    sheet['R9'] = ly_amt_04online # online_amt_ly
    sheet['S9'] = goal_04online # online_goal
    sheet['T9'] = amt_04online  # online_amt
    #HFB05
    sheet['O10'] = amt_05offline[0]  # offline_amt
    sheet['N10'] = goal_05offline[0]  # offline_goal
    sheet['M10'] = ly_amt_05offline[0]  # offline_amt_ly
    sheet['R10'] = ly_amt_05online # online_amt_ly
    sheet['S10'] = goal_05online # online_goal
    sheet['T10'] = amt_05online  # online_amt
    #HFB06
    sheet['O11'] = amt_06offline[0]  # offline_amt
    sheet['N11'] = goal_06offline[0]  # offline_goal
    sheet['M11'] = ly_amt_06offline[0]  # offline_amt_ly
    sheet['R11'] = ly_amt_06online # online_amt_ly
    sheet['S11'] = goal_06online # online_goal
    sheet['T11'] = amt_06online  # online_amt
    #HFB07
    sheet['O12'] = amt_07offline[0]  # offline_amt
    sheet['N12'] = goal_07offline[0]  # offline_goal
    sheet['M12'] = ly_amt_07offline[0]  # offline_amt_ly
    sheet['R12'] = ly_amt_07online # online_amt_ly
    sheet['S12'] = goal_07online # online_goal
    sheet['T12'] = amt_07online  # online_amt
    #HFB08
    sheet['O13'] = amt_08offline[0]  # offline_amt
    sheet['N13'] = goal_08offline[0]  # offline_goal
    sheet['M13'] = ly_amt_08offline[0]  # offline_amt_ly
    sheet['R13'] = ly_amt_08online # online_amt_ly
    sheet['S13'] = goal_08online # online_goal
    sheet['T13'] = amt_08online  # online_amt
    #HFB09
    sheet['O14'] = amt_09offline[0]  # offline_amt
    sheet['N14'] = goal_09offline[0]  # offline_goal
    sheet['M14'] = ly_amt_09offline[0]  # offline_amt_ly
    sheet['R14'] = ly_amt_09online # online_amt_ly
    sheet['S14'] = goal_09online # online_goal
    sheet['T14'] = amt_09online  # online_amt
    #HFB10
    sheet['O15'] = amt_10offline[0]  # offline_amt
    sheet['N15'] = goal_10offline[0]  # offline_goal
    sheet['M15'] = ly_amt_10offline[0]  # offline_amt_ly
    sheet['R15'] = ly_amt_10online # online_amt_ly
    sheet['S15'] = goal_10online # online_goal
    sheet['T15'] = amt_10online  # online_amt
    #HFB11
    sheet['O16'] = amt_11offline[0]  # offline_amt
    sheet['N16'] = goal_11offline[0]  # offline_goal
    sheet['M16'] = ly_amt_11offline[0]  # offline_amt_ly
    sheet['R16'] = ly_amt_11online # online_amt_ly
    sheet['S16'] = goal_11online # online_goal
    sheet['T16'] = amt_11online  # online_amt
    #HFB12
    sheet['O17'] = amt_12offline[0]  # offline_amt
    sheet['N17'] = goal_12offline[0]  # offline_goal
    sheet['M17'] = ly_amt_12offline[0]  # offline_amt_ly
    sheet['R17'] = ly_amt_12online # online_amt_ly
    sheet['S17'] = goal_12online # online_goal
    sheet['T17'] = amt_12online  # online_amt
    #HFB13
    sheet['O18'] = amt_13offline[0]  # offline_amt
    sheet['N18'] = goal_13offline[0]  # offline_goal
    sheet['M18'] = ly_amt_13offline[0]  # offline_amt_ly
    sheet['R18'] = ly_amt_13online # online_amt_ly
    sheet['S18'] = goal_13online # online_goal
    sheet['T18'] = amt_13online  # online_amt
    #HFB14
    sheet['O19'] = amt_14offline[0]  # offline_amt
    sheet['N19'] = goal_14offline[0]  # offline_goal
    sheet['M19'] = ly_amt_14offline[0]  # offline_amt_ly
    sheet['R19'] = ly_amt_14online # online_amt_ly
    sheet['S19'] = goal_14online # online_goal
    sheet['T19'] = amt_14online  # online_amt
    #HFB15
    sheet['O20'] = amt_15offline[0]  # offline_amt
    sheet['N20'] = goal_15offline[0]  # offline_goal
    sheet['M20'] = ly_amt_15offline[0]  # offline_amt_ly
    sheet['R20'] = ly_amt_15online # online_amt_ly
    sheet['S20'] = goal_15online # online_goal
    sheet['T20'] = amt_15online  # online_amt
    #HFB16
    sheet['O21'] = amt_16offline[0]  # offline_amt
    sheet['N21'] = goal_16offline[0]  # offline_goal
    sheet['M21'] = ly_amt_16offline[0]  # offline_amt_ly
    sheet['R21'] = ly_amt_16online # online_amt_ly
    sheet['S21'] = goal_16online # online_goal
    sheet['T21'] = amt_16online  # online_amt
    #HFB17
    sheet['O22'] = amt_17offline[0]  # offline_amt
    sheet['N22'] = goal_17offline[0]  # offline_goal
    sheet['M22'] = ly_amt_17offline[0]  # offline_amt_ly
    sheet['R22'] = ly_amt_17online # online_amt_ly
    sheet['S22'] = goal_17online # online_goal
    sheet['T22'] = amt_17online  # online_amt
    #HFB18
    sheet['O23'] = amt_18offline[0]  # offline_amt
    sheet['N23'] = goal_18offline[0]  # offline_goal
    sheet['M23'] = ly_amt_18offline[0]  # offline_amt_ly
    sheet['R23'] = ly_amt_18online # online_amt_ly
    sheet['S23'] = goal_18online # online_goal
    sheet['T23'] = amt_18online  # online_amt
    #HFB70
    sheet['O24'] = amt_70offline[0]  # offline_amt
    sheet['N24'] = goal_70offline[0]  # offline_goal
    sheet['M24'] = ly_amt_70offline[0]  # offline_amt_ly
    sheet['R24'] = ly_amt_70online # online_amt_ly
    sheet['S24'] = goal_70online # online_goal
    sheet['T24'] = amt_70online  # online_amt
    sheet['P24'] = index_70_offline
    sheet['Q24'] = index_70_offline_ly
    sheet['U24'] = index_70_online
    #HFB95
    sheet['O25'] = amt_95offline[0]  # offline_amt
    sheet['N25'] = goal_95offline[0]  # offline_goal
    sheet['M25'] = ly_amt_95offline[0]  # offline_amt_ly
    sheet['R25'] = ly_amt_95online # online_amt_ly
    sheet['S25'] = goal_95online # online_goal
    sheet['T25'] = amt_95online  # online_amt
    sheet['O26'] = store_offline_amt   # Total
    sheet['T26'] = store_online_amt

    # Store Daily
    sheet['C6'] = store_offline_ly   # offline_store_ly
    sheet['D6'] = store_offline_goal   # offline_store_goal
    sheet['E6'] = store_offline_amt   # offline_store_amt
    sheet['C7'] = store_online_ly   # online_store_ly
    sheet['D7'] = store_online_goal   # online_store_goal
    sheet['E7'] = store_online_amt   # online_store_amt
    sheet['C9'] = store_ikeaFood_ly   # Food_store_ly
    sheet['D9'] = store_ikeaFood_goal   # Food_store_goal
    sheet['E9'] = store_ikeaFood_amt   # Food_store_amt
    sheet['C10'] = store_offline_service_ly   # offline_service_ly
    sheet['D10'] = store_offline_service_goal   # offline_service_goal
    sheet['E10'] = store_offline_service_amt   # offline_service_amt
    sheet['C11'] = store_online_service_ly   # online_service_ly
    sheet['D11'] = store_online_service_goal   # online_service_goal
    sheet['E11'] = store_online_service_amt   # online_service_amt

    #Store_YTD
    sheet['C16'] = store_offline_ly_ytd   #offline_store_ly_ytd
    sheet['D16'] = store_offline_goal_ytd   #offline_store_goal_ytd
    sheet['E16'] = store_offline_amt_ytd   #store_offline_amt_ytd
    sheet['C17'] = store_online_ly_ytd   #online_store_ly_ytd
    sheet['D17'] = store_online_goal_ytd   #online_store_goal_ytd
    sheet['E17'] = store_online_amt_ytd   #online_store_amt_ytd
    sheet['C19'] = store_ikeaFood_ly_ytd   # Food_store_ly_ytd
    sheet['D19'] = store_ikeaFood_goal_ytd   # Food_store_goal_ytd
    sheet['E19'] = store_ikeaFood_amt_ytd   # Food_store_amt_ytd
    sheet['C20'] = store_offline_service_ly_ytd   # offline_service_ly_ytd
    sheet['D20'] = store_offline_service_goal_ytd   # offline_service_goal_ytd
    sheet['E20'] = store_offline_service_amt_ytd   # offline_service_amt_ytd
    sheet['C21'] = store_online_service_ly_ytd   # online_service_ly_ytd
    sheet['D21'] = store_online_service_goal_ytd   # online_service_goal_ytd
    sheet['E21'] = store_online_service_amt_ytd   # online_service_amt_ytd



    # 保存更改
    workbook.save(file_path2)
    print('Data table output-401 processing completed.')
else:
    print("None Data")



