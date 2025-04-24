import pandas as pd
import numpy as np
from openpyxl import load_workbook
from datetime import datetime

# 获取当前日期和时间
now = datetime.now()
formatted_date = now.strftime("%Y-%m-%d")
# 设置店号
store_code = 401

print("Current date:" + formatted_date)

# 读取数据-----------------------------------------------------------------------------------------------------------------------
# 指定文件路径
file_path1 = r'C:\RPAData\api_result.csv'
file_path2 = r'C:\RPAData\output-401-2.xlsx'

# 读取 CSV 文件
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
            first_valid_date = excel1_df['agg_date'].dropna().iloc[0]

            # 格式化日期
            formatted_first_agg_date = first_valid_date.strftime("%Y-%m-%d")
            print("Formatted first agg_date:", formatted_first_agg_date)
        else:
            print("No valid dates found in 'agg_date'.")

    except Exception as e:
        print("An error occurred while processing dates:", e)
else:
    print("DataFrame is empty or 'agg_date' column does not exist.")


#offline store
filtered_df_offline_store = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (excel1_df['product_type'] == 'goods') & (excel1_df['product_type'] != 'other')].copy()
#offline store_tyd
filtered_df_offline_store_ytd = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (excel1_df['product_type'] == 'goods') & (excel1_df['product_type'] != 'other')].copy()
#Online_store
filtered_df_online_store = excel1_df[(excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (excel1_df['product_type'] == 'goods') & (excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()


# HFB01 offline
filtered_df_01offline = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (
                excel1_df['hfb_no'] == 1)].copy()
# HFB01 online
filtered_df_01online = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (
                excel1_df['hfb_no'] == 1) & (excel1_df['product_type'] == 'goods') & (
                excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()

# HFB02 offline
filtered_df_02offline = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (
                excel1_df['hfb_no'] == 2)].copy()
# HFB02 online
filtered_df_02online = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (
                excel1_df['hfb_no'] == 2) & (excel1_df['product_type'] == 'goods') & (
                excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()

# HFB03 offline
filtered_df_03offline = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (
                excel1_df['hfb_no'] == 3)].copy()
# HFB03 online
filtered_df_03online = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (
                excel1_df['hfb_no'] == 3) & (excel1_df['product_type'] == 'goods') & (
                excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()

# HFB04 offline
filtered_df_04offline = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (
                excel1_df['hfb_no'] == 4)].copy()
# HFB04 online
filtered_df_04online = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (
                excel1_df['hfb_no'] == 4) & (excel1_df['product_type'] == 'goods') & (
                excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()

# HFB05 offline
filtered_df_05offline = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (
                excel1_df['hfb_no'] == 5)].copy()
# HFB05 online
filtered_df_05online = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (
                excel1_df['hfb_no'] == 5) & (excel1_df['product_type'] == 'goods') & (
                excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()

# HFB06 offline
filtered_df_06offline = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (
                excel1_df['hfb_no'] == 6)].copy()
# HFB06 online
filtered_df_06online = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (
                excel1_df['hfb_no'] == 6) & (excel1_df['product_type'] == 'goods') & (
                excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()

# HFB07 offline
filtered_df_07offline = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (
                excel1_df['hfb_no'] == 7)].copy()
# HFB07 online
filtered_df_07online = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (
                excel1_df['hfb_no'] == 7) & (excel1_df['product_type'] == 'goods') & (
                excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()

# HFB08 offline
filtered_df_08offline = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (
                excel1_df['hfb_no'] == 8)].copy()
# HFB08 online
filtered_df_08online = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (
                excel1_df['hfb_no'] == 8) & (excel1_df['product_type'] == 'goods') & (
                excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()

# HFB09 offline
filtered_df_09offline = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (
                excel1_df['hfb_no'] == 9)].copy()
# HFB09 online
filtered_df_09online = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (
                excel1_df['hfb_no'] == 9) & (excel1_df['product_type'] == 'goods') & (
                excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()

# HFB10 offline
filtered_df_10offline = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (
                excel1_df['hfb_no'] == 10)].copy()
# HFB10 online
filtered_df_10online = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (
                excel1_df['hfb_no'] == 10) & (excel1_df['product_type'] == 'goods') & (
                excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()

# HFB11 offline
filtered_df_11offline = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (
                excel1_df['hfb_no'] == 11)].copy()
# HFB11 online
filtered_df_11online = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (
                excel1_df['hfb_no'] == 11) & (excel1_df['product_type'] == 'goods') & (
                excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()

# HFB12 offline
filtered_df_12offline = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (
                excel1_df['hfb_no'] == 12)].copy()
# HFB12 online
filtered_df_12online = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (
                excel1_df['hfb_no'] == 12) & (excel1_df['product_type'] == 'goods') & (
                excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()

# HFB13 offline
filtered_df_13offline = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (
                excel1_df['hfb_no'] == 13)].copy()
# HFB13 online
filtered_df_13online = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (
                excel1_df['hfb_no'] == 13) & (excel1_df['product_type'] == 'goods') & (
                excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()

# HFB14 offline
filtered_df_14offline = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (
                excel1_df['hfb_no'] == 14)].copy()
# HFB14 online
filtered_df_14online = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (
                excel1_df['hfb_no'] == 14) & (excel1_df['product_type'] == 'goods') & (
                excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()

# HFB15 offline
filtered_df_15offline = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (
                excel1_df['hfb_no'] == 15)].copy()
# HFB15 online
filtered_df_15online = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (
                excel1_df['hfb_no'] == 15) & (excel1_df['product_type'] == 'goods') & (
                excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()

# HFB16 offline
filtered_df_16offline = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (
                excel1_df['hfb_no'] == 16)].copy()
# HFB16 online
filtered_df_16online = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (
                excel1_df['hfb_no'] == 16) & (excel1_df['product_type'] == 'goods') & (
                excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()

# HFB17 offline
filtered_df_17offline = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (
                excel1_df['hfb_no'] == 17)].copy()
# HFB17 online
filtered_df_17online = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (
                excel1_df['hfb_no'] == 17) & (excel1_df['product_type'] == 'goods') & (
                excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()

# HFB18 offline
filtered_df_18offline = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (
                excel1_df['hfb_no'] == 18)].copy()
# HFB18 online
filtered_df_18online = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (
                excel1_df['hfb_no'] == 18) & (excel1_df['product_type'] == 'goods') & (
                excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()

# HFB70 offline
filtered_df_70offline = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'OFFLINE') & (
                excel1_df['hfb_no'] == 70)].copy()
# HFB18 online
filtered_df_70online = excel1_df[
    (excel1_df['market_code'] == store_code) & (excel1_df['sales_channel_lvl1'] == 'ONLINE') & (
                excel1_df['hfb_no'] == 70) & (excel1_df['product_type'] == 'goods') & (
                excel1_df['sales_channel'] != 'NEW PLATFORMS')].copy()

# 取值并取整-------------------------------------------------------------------------------------------------------------------------------
# HFB01 offline
amt_01offline = np.round(filtered_df_01offline['sale_net_amt_fytd'].values / 1000).astype(int)
ly_amt_01offline = np.round(filtered_df_01offline['ly_sale_net_amt_fytd'].values / 1000).astype(int)
goal_01offline = np.round(filtered_df_01offline['sales_goal_fytd'].values / 1000).astype(int)
index_01_offline = np.round(filtered_df_01offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_01offline['sales_goal_fytd'].values) * 100
index_ly_01_offline = np.round(filtered_df_01offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_01offline['ly_sale_net_amt_fytd'].values) * 100

# HFB01 online
amt_01online = np.round(filtered_df_01online['sale_net_amt_fytd'].sum() / 1000).astype(int)
goal_01online = np.round(filtered_df_01online['sales_goal_fytd'].sum() / 1000).astype(int)
ly_amt_01online = np.round(filtered_df_01online['ly_sale_net_amt_fytd'].sum() / 1000).astype(int)
index_01_online = np.round(filtered_df_01online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_01online['sales_goal_fytd'].sum()) * 100
index_ly_01_online = np.round(filtered_df_01online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_01online['ly_sale_net_amt_fytd'].sum()) * 100

# HFB02 offline
amt_02offline = np.round(filtered_df_02offline['sale_net_amt_fytd'].values / 1000).astype(int)
ly_amt_02offline = np.round(filtered_df_02offline['ly_sale_net_amt_fytd'].values / 1000).astype(int)
goal_02offline = np.round(filtered_df_02offline['sales_goal_fytd'].values / 1000).astype(int)
index_02_offline = np.round(filtered_df_02offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_02offline['sales_goal_fytd'].values) * 100
index_ly_02_offline = np.round(filtered_df_02offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_02offline['ly_sale_net_amt_fytd'].values) * 100

# HFB02 online
amt_02online = np.round(filtered_df_02online['sale_net_amt_fytd'].sum() / 1000).astype(int)
goal_02online = np.round(filtered_df_02online['sales_goal_fytd'].sum() / 1000).astype(int)
ly_amt_02online = np.round(filtered_df_02online['ly_sale_net_amt_fytd'].sum() / 1000).astype(int)
index_02_online = np.round(filtered_df_02online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_02online['sales_goal_fytd'].sum()) * 100
index_ly_02_online = np.round(filtered_df_02online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_02online['ly_sale_net_amt_fytd'].sum()) * 100

# HFB03 offline
amt_03offline = np.round(filtered_df_03offline['sale_net_amt_fytd'].values / 1000).astype(int)
ly_amt_03offline = np.round(filtered_df_03offline['ly_sale_net_amt_fytd'].values / 1000).astype(int)
goal_03offline = np.round(filtered_df_03offline['sales_goal_fytd'].values / 1000).astype(int)
index_03_offline = np.round(filtered_df_03offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_03offline['sales_goal_fytd'].values) * 100
index_ly_03_offline = np.round(filtered_df_03offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_03offline['ly_sale_net_amt_fytd'].values) * 100

# HFB03 online
amt_03online = np.round(filtered_df_03online['sale_net_amt_fytd'].sum() / 1000).astype(int)
goal_03online = np.round(filtered_df_03online['sales_goal_fytd'].sum() / 1000).astype(int)
ly_amt_03online = np.round(filtered_df_03online['ly_sale_net_amt_fytd'].sum() / 1000).astype(int)
index_03_online = np.round(filtered_df_03online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_03online['sales_goal_fytd'].sum()) * 100
index_ly_03_online = np.round(filtered_df_03online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_03online['ly_sale_net_amt_fytd'].sum()) * 100

# HFB04 offline
amt_04offline = np.round(filtered_df_04offline['sale_net_amt_fytd'].values / 1000).astype(int)
ly_amt_04offline = np.round(filtered_df_04offline['ly_sale_net_amt_fytd'].values / 1000).astype(int)
goal_04offline = np.round(filtered_df_04offline['sales_goal_fytd'].values / 1000).astype(int)
index_04_offline = np.round(filtered_df_04offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_04offline['sales_goal_fytd'].values) * 100
index_ly_04_offline = np.round(filtered_df_04offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_04offline['ly_sale_net_amt_fytd'].values) * 100

# HFB04 online
amt_04online = np.round(filtered_df_04online['sale_net_amt_fytd'].sum() / 1000).astype(int)
goal_04online = np.round(filtered_df_04online['sales_goal_fytd'].sum() / 1000).astype(int)
ly_amt_04online = np.round(filtered_df_04online['ly_sale_net_amt_fytd'].sum() / 1000).astype(int)
index_04_online = np.round(filtered_df_04online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_04online['sales_goal_fytd'].sum()) * 100
index_ly_04_online = np.round(filtered_df_04online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_04online['ly_sale_net_amt_fytd'].sum()) * 100

# HFB05 offline
amt_05offline = np.round(filtered_df_05offline['sale_net_amt_fytd'].values / 1000).astype(int)
ly_amt_05offline = np.round(filtered_df_05offline['ly_sale_net_amt_fytd'].values / 1000).astype(int)
goal_05offline = np.round(filtered_df_05offline['sales_goal_fytd'].values / 1000).astype(int)
index_05_offline = np.round(filtered_df_05offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_05offline['sales_goal_fytd'].values) * 100
index_ly_05_offline = np.round(filtered_df_05offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_05offline['ly_sale_net_amt_fytd'].values) * 100

# HFB05 online
amt_05online = np.round(filtered_df_05online['sale_net_amt_fytd'].sum() / 1000).astype(int)
goal_05online = np.round(filtered_df_05online['sales_goal_fytd'].sum() / 1000).astype(int)
ly_amt_05online = np.round(filtered_df_05online['ly_sale_net_amt_fytd'].sum() / 1000).astype(int)
index_05_online = np.round(filtered_df_05online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_05online['sales_goal_fytd'].sum()) * 100
index_ly_05_online = np.round(filtered_df_05online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_05online['ly_sale_net_amt_fytd'].sum()) * 100

# HFB06 offline
amt_06offline = np.round(filtered_df_06offline['sale_net_amt_fytd'].values / 1000).astype(int)
ly_amt_06offline = np.round(filtered_df_06offline['ly_sale_net_amt_fytd'].values / 1000).astype(int)
goal_06offline = np.round(filtered_df_06offline['sales_goal_fytd'].values / 1000).astype(int)
index_06_offline = np.round(filtered_df_06offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_06offline['sales_goal_fytd'].values) * 100
index_ly_06_offline = np.round(filtered_df_06offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_06offline['ly_sale_net_amt_fytd'].values) * 100

# HFB06 online
amt_06online = np.round(filtered_df_06online['sale_net_amt_fytd'].sum() / 1000).astype(int)
goal_06online = np.round(filtered_df_06online['sales_goal_fytd'].sum() / 1000).astype(int)
ly_amt_06online = np.round(filtered_df_06online['ly_sale_net_amt_fytd'].sum() / 1000).astype(int)
index_06_online = np.round(filtered_df_06online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_06online['sales_goal_fytd'].sum()) * 100
index_ly_06_online = np.round(filtered_df_06online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_06online['ly_sale_net_amt_fytd'].sum()) * 100

# HFB07 offline
amt_07offline = np.round(filtered_df_07offline['sale_net_amt_fytd'].values / 1000).astype(int)
ly_amt_07offline = np.round(filtered_df_07offline['ly_sale_net_amt_fytd'].values / 1000).astype(int)
goal_07offline = np.round(filtered_df_07offline['sales_goal_fytd'].values / 1000).astype(int)
index_07_offline = np.round(filtered_df_07offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_07offline['sales_goal_fytd'].values) * 100
index_ly_07_offline = np.round(filtered_df_07offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_07offline['ly_sale_net_amt_fytd'].values) * 100

# HFB07 online
amt_07online = np.round(filtered_df_07online['sale_net_amt_fytd'].sum() / 1000).astype(int)
goal_07online = np.round(filtered_df_07online['sales_goal_fytd'].sum() / 1000).astype(int)
ly_amt_07online = np.round(filtered_df_07online['ly_sale_net_amt_fytd'].sum() / 1000).astype(int)
index_07_online = np.round(filtered_df_07online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_07online['sales_goal_fytd'].sum()) * 100
index_ly_07_online = np.round(filtered_df_07online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_07online['ly_sale_net_amt_fytd'].sum()) * 100

# HFB08 offline
amt_08offline = np.round(filtered_df_08offline['sale_net_amt_fytd'].values / 1000).astype(int)
ly_amt_08offline = np.round(filtered_df_08offline['ly_sale_net_amt_fytd'].values / 1000).astype(int)
goal_08offline = np.round(filtered_df_08offline['sales_goal_fytd'].values / 1000).astype(int)
index_08_offline = np.round(filtered_df_08offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_08offline['sales_goal_fytd'].values) * 100
index_ly_08_offline = np.round(filtered_df_08offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_08offline['ly_sale_net_amt_fytd'].values) * 100

# HFB08 online
amt_08online = np.round(filtered_df_08online['sale_net_amt_fytd'].sum() / 1000).astype(int)
goal_08online = np.round(filtered_df_08online['sales_goal_fytd'].sum() / 1000).astype(int)
ly_amt_08online = np.round(filtered_df_08online['ly_sale_net_amt_fytd'].sum() / 1000).astype(int)
index_08_online = np.round(filtered_df_08online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_08online['sales_goal_fytd'].sum()) * 100
index_ly_08_online = np.round(filtered_df_08online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_08online['ly_sale_net_amt_fytd'].sum()) * 100

# HFB09 offline
amt_09offline = np.round(filtered_df_09offline['sale_net_amt_fytd'].values / 1000).astype(int)
ly_amt_09offline = np.round(filtered_df_09offline['ly_sale_net_amt_fytd'].values / 1000).astype(int)
goal_09offline = np.round(filtered_df_09offline['sales_goal_fytd'].values / 1000).astype(int)
index_09_offline = np.round(filtered_df_09offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_09offline['sales_goal_fytd'].values) * 100
index_ly_09_offline = np.round(filtered_df_09offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_09offline['ly_sale_net_amt_fytd'].values) * 100

# HFB09 online
amt_09online = np.round(filtered_df_09online['sale_net_amt_fytd'].sum() / 1000).astype(int)
goal_09online = np.round(filtered_df_09online['sales_goal_fytd'].sum() / 1000).astype(int)
ly_amt_09online = np.round(filtered_df_09online['ly_sale_net_amt_fytd'].sum() / 1000).astype(int)
index_09_online = np.round(filtered_df_09online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_09online['sales_goal_fytd'].sum()) * 100
index_ly_09_online = np.round(filtered_df_09online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_09online['ly_sale_net_amt_fytd'].sum()) * 100

# HFB10 offline
amt_10offline = np.round(filtered_df_10offline['sale_net_amt_fytd'].values / 1000).astype(int)
ly_amt_10offline = np.round(filtered_df_10offline['ly_sale_net_amt_fytd'].values / 1000).astype(int)
goal_10offline = np.round(filtered_df_10offline['sales_goal_fytd'].values / 1000).astype(int)
index_10_offline = np.round(filtered_df_10offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_10offline['sales_goal_fytd'].values) * 100
index_ly_10_offline = np.round(filtered_df_10offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_10offline['ly_sale_net_amt_fytd'].values) * 100

# HFB10 online
amt_10online = np.round(filtered_df_10online['sale_net_amt_fytd'].sum() / 1000).astype(int)
goal_10online = np.round(filtered_df_10online['sales_goal_fytd'].sum() / 1000).astype(int)
ly_amt_10online = np.round(filtered_df_10online['ly_sale_net_amt_fytd'].sum() / 1000).astype(int)
index_10_online = np.round(filtered_df_10online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_10online['sales_goal_fytd'].sum()) * 100
index_ly_10_online = np.round(filtered_df_10online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_10online['ly_sale_net_amt_fytd'].sum()) * 100

# HFB11 offline
amt_11offline = np.round(filtered_df_11offline['sale_net_amt_fytd'].values / 1000).astype(int)
ly_amt_11offline = np.round(filtered_df_11offline['ly_sale_net_amt_fytd'].values / 1000).astype(int)
goal_11offline = np.round(filtered_df_11offline['sales_goal_fytd'].values / 1000).astype(int)
index_11_offline = np.round(filtered_df_11offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_11offline['sales_goal_fytd'].values) * 100
index_ly_11_offline = np.round(filtered_df_11offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_11offline['ly_sale_net_amt_fytd'].values) * 100

# HFB11 online
amt_11online = np.round(filtered_df_11online['sale_net_amt_fytd'].sum() / 1000).astype(int)
goal_11online = np.round(filtered_df_11online['sales_goal_fytd'].sum() / 1000).astype(int)
ly_amt_11online = np.round(filtered_df_11online['ly_sale_net_amt_fytd'].sum() / 1000).astype(int)
index_11_online = np.round(filtered_df_11online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_11online['sales_goal_fytd'].sum()) * 100
index_ly_11_online = np.round(filtered_df_11online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_11online['ly_sale_net_amt_fytd'].sum()) * 100

# HFB12 offline
amt_12offline = np.round(filtered_df_12offline['sale_net_amt_fytd'].values / 1000).astype(int)
ly_amt_12offline = np.round(filtered_df_12offline['ly_sale_net_amt_fytd'].values / 1000).astype(int)
goal_12offline = np.round(filtered_df_12offline['sales_goal_fytd'].values / 1000).astype(int)
index_12_offline = np.round(filtered_df_12offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_12offline['sales_goal_fytd'].values) * 100
index_ly_12_offline = np.round(filtered_df_12offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_12offline['ly_sale_net_amt_fytd'].values) * 100

# HFB12 online
amt_12online = np.round(filtered_df_12online['sale_net_amt_fytd'].sum() / 1000).astype(int)
goal_12online = np.round(filtered_df_12online['sales_goal_fytd'].sum() / 1000).astype(int)
ly_amt_12online = np.round(filtered_df_12online['ly_sale_net_amt_fytd'].sum() / 1000).astype(int)
index_12_online = np.round(filtered_df_12online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_12online['sales_goal_fytd'].sum()) * 100
index_ly_12_online = np.round(filtered_df_12online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_12online['ly_sale_net_amt_fytd'].sum()) * 100

# HFB13 offline
amt_13offline = np.round(filtered_df_13offline['sale_net_amt_fytd'].values / 1000).astype(int)
ly_amt_13offline = np.round(filtered_df_13offline['ly_sale_net_amt_fytd'].values / 1000).astype(int)
goal_13offline = np.round(filtered_df_13offline['sales_goal_fytd'].values / 1000).astype(int)
index_13_offline = np.round(filtered_df_13offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_13offline['sales_goal_fytd'].values) * 100
index_ly_13_offline = np.round(filtered_df_13offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_13offline['ly_sale_net_amt_fytd'].values) * 100

# HFB13 online
amt_13online = np.round(filtered_df_13online['sale_net_amt_fytd'].sum() / 1000).astype(int)
goal_13online = np.round(filtered_df_13online['sales_goal_fytd'].sum() / 1000).astype(int)
ly_amt_13online = np.round(filtered_df_13online['ly_sale_net_amt_fytd'].sum() / 1000).astype(int)
index_13_online = np.round(filtered_df_13online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_13online['sales_goal_fytd'].sum()) * 100
index_ly_13_online = np.round(filtered_df_13online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_13online['ly_sale_net_amt_fytd'].sum()) * 100

# HFB14 offline
amt_14offline = np.round(filtered_df_14offline['sale_net_amt_fytd'].values / 1000).astype(int)
ly_amt_14offline = np.round(filtered_df_14offline['ly_sale_net_amt_fytd'].values / 1000).astype(int)
goal_14offline = np.round(filtered_df_14offline['sales_goal_fytd'].values / 1000).astype(int)
index_14_offline = np.round(filtered_df_14offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_14offline['sales_goal_fytd'].values) * 100
index_ly_14_offline = np.round(filtered_df_14offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_14offline['ly_sale_net_amt_fytd'].values) * 100

# HFB14 online
amt_14online = np.round(filtered_df_14online['sale_net_amt_fytd'].sum() / 1000).astype(int)
goal_14online = np.round(filtered_df_14online['sales_goal_fytd'].sum() / 1000).astype(int)
ly_amt_14online = np.round(filtered_df_14online['ly_sale_net_amt_fytd'].sum() / 1000).astype(int)
index_14_online = np.round(filtered_df_14online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_14online['sales_goal_fytd'].sum()) * 100
index_ly_14_online = np.round(filtered_df_14online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_14online['ly_sale_net_amt_fytd'].sum()) * 100

# HFB15 offline
amt_15offline = np.round(filtered_df_15offline['sale_net_amt_fytd'].values / 1000).astype(int)
ly_amt_15offline = np.round(filtered_df_15offline['ly_sale_net_amt_fytd'].values / 1000).astype(int)
goal_15offline = np.round(filtered_df_15offline['sales_goal_fytd'].values / 1000).astype(int)
index_15_offline = np.round(filtered_df_15offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_15offline['sales_goal_fytd'].values) * 100
index_ly_15_offline = np.round(filtered_df_15offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_15offline['ly_sale_net_amt_fytd'].values) * 100

# HFB15 online
amt_15online = np.round(filtered_df_15online['sale_net_amt_fytd'].sum() / 1000).astype(int)
goal_15online = np.round(filtered_df_15online['sales_goal_fytd'].sum() / 1000).astype(int)
ly_amt_15online = np.round(filtered_df_15online['ly_sale_net_amt_fytd'].sum() / 1000).astype(int)
index_15_online = np.round(filtered_df_15online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_15online['sales_goal_fytd'].sum()) * 100
index_ly_15_online = np.round(filtered_df_15online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_15online['ly_sale_net_amt_fytd'].sum()) * 100

# HFB16 offline
amt_16offline = np.round(filtered_df_16offline['sale_net_amt_fytd'].values / 1000).astype(int)
ly_amt_16offline = np.round(filtered_df_16offline['ly_sale_net_amt_fytd'].values / 1000).astype(int)
goal_16offline = np.round(filtered_df_16offline['sales_goal_fytd'].values / 1000).astype(int)
index_16_offline = np.round(filtered_df_16offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_16offline['sales_goal_fytd'].values) * 100
index_ly_16_offline = np.round(filtered_df_16offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_16offline['ly_sale_net_amt_fytd'].values) * 100

# HFB16 online
amt_16online = np.round(filtered_df_16online['sale_net_amt_fytd'].sum() / 1000).astype(int)
goal_16online = np.round(filtered_df_16online['sales_goal_fytd'].sum() / 1000).astype(int)
ly_amt_16online = np.round(filtered_df_16online['ly_sale_net_amt_fytd'].sum() / 1000).astype(int)
index_16_online = np.round(filtered_df_16online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_16online['sales_goal_fytd'].sum()) * 100
index_ly_16_online = np.round(filtered_df_16online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_16online['ly_sale_net_amt_fytd'].sum()) * 100

# HFB17 offline
amt_17offline = np.round(filtered_df_17offline['sale_net_amt_fytd'].values / 1000).astype(int)
ly_amt_17offline = np.round(filtered_df_17offline['ly_sale_net_amt_fytd'].values / 1000).astype(int)
goal_17offline = np.round(filtered_df_17offline['sales_goal_fytd'].values / 1000).astype(int)
index_17_offline = np.round(filtered_df_17offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_17offline['sales_goal_fytd'].values) * 100
index_ly_17_offline = np.round(filtered_df_17offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_17offline['ly_sale_net_amt_fytd'].values) * 100

# HFB17 online
amt_17online = np.round(filtered_df_17online['sale_net_amt_fytd'].sum() / 1000).astype(int)
goal_17online = np.round(filtered_df_17online['sales_goal_fytd'].sum() / 1000).astype(int)
ly_amt_17online = np.round(filtered_df_17online['ly_sale_net_amt_fytd'].sum() / 1000).astype(int)
index_17_online = np.round(filtered_df_17online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_17online['sales_goal_fytd'].sum()) * 100
index_ly_17_online = np.round(filtered_df_17online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_17online['ly_sale_net_amt_fytd'].sum()) * 100

# HFB18 offline
amt_18offline = np.round(filtered_df_18offline['sale_net_amt_fytd'].values / 1000).astype(int)
ly_amt_18offline = np.round(filtered_df_18offline['ly_sale_net_amt_fytd'].values / 1000).astype(int)
goal_18offline = np.round(filtered_df_18offline['sales_goal_fytd'].values / 1000).astype(int)
index_18_offline = np.round(filtered_df_18offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_18offline['sales_goal_fytd'].values) * 100
index_ly_18_offline = np.round(filtered_df_18offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_18offline['ly_sale_net_amt_fytd'].values) * 100

# HFB18 online
amt_18online = np.round(filtered_df_18online['sale_net_amt_fytd'].sum() / 1000).astype(int)
goal_18online = np.round(filtered_df_18online['sales_goal_fytd'].sum() / 1000).astype(int)
ly_amt_18online = np.round(filtered_df_18online['ly_sale_net_amt_fytd'].sum() / 1000).astype(int)
index_18_online = np.round(filtered_df_18online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_18online['sales_goal_fytd'].sum()) *100
index_ly_18_online = np.round(filtered_df_18online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_18online['ly_sale_net_amt_fytd'].sum()) * 100

# HFB70 offline
amt_70offline = np.round(filtered_df_70offline['sale_net_amt_fytd'].values / 1000).astype(int)
ly_amt_70offline = np.round(filtered_df_70offline['ly_sale_net_amt_fytd'].values / 1000).astype(int)
goal_70offline = np.round(filtered_df_70offline['sales_goal_fytd'].values / 1000).astype(int)
index_70_offline = np.round(filtered_df_70offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_70offline['sales_goal_fytd'].values) * 100
index_ly_70_offline = np.round(filtered_df_70offline['sale_net_amt_fytd'].values) / np.round(
    filtered_df_70offline['ly_sale_net_amt_fytd'].values) * 100

# HFB70 online
amt_70online = np.round(filtered_df_70online['sale_net_amt_fytd'].sum() / 1000).astype(int)
goal_70online = np.round(filtered_df_70online['sales_goal_fytd'].sum() / 1000).astype(int)
ly_amt_70online = np.round(filtered_df_70online['ly_sale_net_amt_fytd'].sum() / 1000).astype(int)
index_70_online = np.round(filtered_df_70online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_70online['sales_goal_fytd'].sum()) *100
index_ly_70_online = np.round(filtered_df_70online['sale_net_amt_fytd'].sum()) / np.round(
    filtered_df_70online['ly_sale_net_amt_fytd'].sum()) * 100


store_offline_ly = np.round(filtered_df_offline_store['ly_sale_net_amt_fytd'].sum() / 1000).astype(int)
store_offline_goal = np.round(filtered_df_offline_store['sales_goal_fytd'].sum() / 1000).astype(int)
store_offline_amt = np.round(filtered_df_offline_store['sale_net_amt_fytd'].sum() / 1000).astype(int)
store_online_ly = np.round(filtered_df_online_store['ly_sale_net_amt_fytd'].sum() / 1000).astype(int)
store_online_goal = np.round(filtered_df_online_store['sales_goal_fytd'].sum() / 1000).astype(int)
store_online_amt = np.round(filtered_df_online_store['sale_net_amt_fytd'].sum() / 1000).astype(int)

# 写入数据----------------------------------------------------------------------------------------------------------
# 加载现有的 Excel 文件
workbook = load_workbook(file_path2)
# 选择要写入的工作表
sheet = workbook['HZ 日结模版']
excel_df = pd.read_excel(file_path2, sheet_name='HZ 日结模版')

# 写入单元格
if len(ly_amt_01offline) > 0 and len(goal_01offline) > 0:
    sheet['C5'] = formatted_first_agg_date  # 写入当天日期
    #Total
    sheet['B27'] = store_offline_ly
    sheet['C27'] = store_offline_goal
    sheet['D27'] = store_offline_amt
    sheet['G27'] = store_online_ly
    sheet['H27'] = store_online_goal
    sheet['I27'] = store_online_amt



    # HFB01 offline
    sheet['B7'] = ly_amt_01offline[0]
    sheet['C7'] = goal_01offline[0]
    sheet['D7'] = amt_01offline[0]
    # HFB01 online
    sheet['G7'] = ly_amt_01online  # online_amt_ly
    sheet['H7'] = goal_01online  # online_goal
    sheet['I7'] = amt_01online  # online_amt
    # HFB01_Index
    sheet['E7'] = index_01_offline[0]
    sheet['F7'] = index_ly_01_offline[0]
    sheet['J7'] = index_01_online
    sheet['K7'] = index_ly_01_online

# HFB02
if len(ly_amt_02offline) > 0 and len(goal_02offline) > 0:
    sheet['B8'] = ly_amt_02offline[0]
    sheet['C8'] = goal_02offline[0]
    sheet['D8'] = amt_02offline[0]
    sheet['G8'] = ly_amt_02online  # online_amt_ly
    sheet['H8'] = goal_02online  # online_goal
    sheet['I8'] = amt_02online  # online_amt
    # HFB02_Index
    sheet['E8'] = index_02_offline[0]
    sheet['F8'] = index_ly_02_offline[0]
    sheet['J8'] = index_02_online
    sheet['K8'] = index_ly_02_online

# HFB03
if len(ly_amt_03offline) > 0 and len(goal_03offline) > 0:
    sheet['B9'] = ly_amt_03offline[0]
    sheet['C9'] = goal_03offline[0]
    sheet['D9'] = amt_03offline[0]
    sheet['G9'] = ly_amt_03online  # online_amt_ly
    sheet['H9'] = goal_03online  # online_goal
    sheet['I9'] = amt_03online  # online_amt
    # HFB03_Index
    sheet['E9'] = index_03_offline[0]
    sheet['F9'] = index_ly_03_offline[0]
    sheet['J9'] = index_03_online
    sheet['K9'] = index_ly_03_online

# HFB04
if len(ly_amt_04offline) > 0 and len(goal_04offline) > 0:
    sheet['B10'] = ly_amt_04offline[0]
    sheet['C10'] = goal_04offline[0]
    sheet['D10'] = amt_04offline[0]
    sheet['G10'] = ly_amt_04online  # online_amt_ly
    sheet['H10'] = goal_04online  # online_goal
    sheet['I10'] = amt_04online  # online_amt
    # HFB04_Index
    sheet['E10'] = index_04_offline[0]
    sheet['F10'] = index_ly_04_offline[0]
    sheet['J10'] = index_04_online
    sheet['K10'] = index_ly_04_online

# HFB05
if len(ly_amt_05offline) > 0 and len(goal_05offline) > 0:
    sheet['B11'] = ly_amt_05offline[0]
    sheet['C11'] = goal_05offline[0]
    sheet['D11'] = amt_05offline[0]
    sheet['G11'] = ly_amt_05online  # online_amt_ly
    sheet['H11'] = goal_05online  # online_goal
    sheet['I11'] = amt_05online  # online_amt
    # HFB05_Index
    sheet['E11'] = index_05_offline[0]
    sheet['F11'] = index_ly_05_offline[0]
    sheet['J11'] = index_05_online
    sheet['K11'] = index_ly_05_online

# HFB06
if len(ly_amt_06offline) > 0 and len(goal_06offline) > 0:
    sheet['B12'] = ly_amt_06offline[0]
    sheet['C12'] = goal_06offline[0]
    sheet['D12'] = amt_06offline[0]
    sheet['G12'] = ly_amt_06online  # online_amt_ly
    sheet['H12'] = goal_06online  # online_goal
    sheet['I12'] = amt_06online  # online_amt
    # HFB06_Index
    sheet['E12'] = index_06_offline[0]
    sheet['F12'] = index_ly_06_offline[0]
    sheet['J12'] = index_06_online
    sheet['K12'] = index_ly_06_online

# HFB07
if len(ly_amt_07offline) > 0 and len(goal_07offline) > 0:
    sheet['B13'] = ly_amt_07offline[0]
    sheet['C13'] = goal_07offline[0]
    sheet['D13'] = amt_07offline[0]
    sheet['G13'] = ly_amt_07online  # online_amt_ly
    sheet['H13'] = goal_07online  # online_goal
    sheet['I13'] = amt_07online  # online_amt
    # HFB07_Index
    sheet['E13'] = index_07_offline[0]
    sheet['F13'] = index_ly_07_offline[0]
    sheet['J13'] = index_07_online
    sheet['K13'] = index_ly_07_online

# HFB08
if len(ly_amt_08offline) > 0 and len(goal_08offline) > 0:
    sheet['B14'] = ly_amt_08offline[0]
    sheet['C14'] = goal_08offline[0]
    sheet['D14'] = amt_08offline[0]
    sheet['G14'] = ly_amt_08online  # online_amt_ly
    sheet['H14'] = goal_08online  # online_goal
    sheet['I14'] = amt_08online  # online_amt
    # HFB08_Index
    sheet['E14'] = index_08_offline[0]
    sheet['F14'] = index_ly_08_offline[0]
    sheet['J14'] = index_08_online
    sheet['K14'] = index_ly_08_online

# HFB09
if len(ly_amt_09offline) > 0 and len(goal_09offline) > 0:
    sheet['B15'] = ly_amt_09offline[0]
    sheet['C15'] = goal_09offline[0]
    sheet['D15'] = amt_09offline[0]
    sheet['G15'] = ly_amt_09online  # online_amt_ly
    sheet['H15'] = goal_09online  # online_goal
    sheet['I15'] = amt_09online  # online_amt
    # HFB09_Index
    sheet['E15'] = index_09_offline[0]
    sheet['F15'] = index_ly_09_offline[0]
    sheet['J15'] = index_09_online
    sheet['K15'] = index_ly_09_online

# HFB10
if len(ly_amt_10offline) > 0 and len(goal_10offline) > 0:
    sheet['B16'] = ly_amt_10offline[0]
    sheet['C16'] = goal_10offline[0]
    sheet['D16'] = amt_10offline[0]
    sheet['G16'] = ly_amt_10online  # online_amt_ly
    sheet['H16'] = goal_10online  # online_goal
    sheet['I16'] = amt_10online  # online_amt
    # HFB10_Index
    sheet['E16'] = index_10_offline[0]
    sheet['F16'] = index_ly_10_offline[0]
    sheet['J16'] = index_10_online
    sheet['K16'] = index_ly_10_online

# HFB11
if len(ly_amt_11offline) > 0 and len(goal_11offline) > 0:
    sheet['B17'] = ly_amt_11offline[0]
    sheet['C17'] = goal_11offline[0]
    sheet['D17'] = amt_11offline[0]
    sheet['G17'] = ly_amt_11online  # online_amt_ly
    sheet['H17'] = goal_11online  # online_goal
    sheet['I17'] = amt_11online  # online_amt
    # HFB11_Index
    sheet['E17'] = index_11_offline[0]
    sheet['F17'] = index_ly_11_offline[0]
    sheet['J17'] = index_11_online
    sheet['K17'] = index_ly_11_online

# HFB12
if len(ly_amt_12offline) > 0 and len(goal_12offline) > 0:
    sheet['B18'] = ly_amt_12offline[0]
    sheet['C18'] = goal_12offline[0]
    sheet['D18'] = amt_12offline[0]
    sheet['G18'] = ly_amt_12online  # online_amt_ly
    sheet['H18'] = goal_12online  # online_goal
    sheet['I18'] = amt_12online  # online_amt
    # HFB12_Index
    sheet['E18'] = index_12_offline[0]
    sheet['F18'] = index_ly_12_offline[0]
    sheet['J18'] = index_12_online
    sheet['K18'] = index_ly_12_online

# HFB13
if len(ly_amt_13offline) > 0 and len(goal_13offline) > 0:
    sheet['B19'] = ly_amt_13offline[0]
    sheet['C19'] = goal_13offline[0]
    sheet['D19'] = amt_13offline[0]
    sheet['G19'] = ly_amt_13online  # online_amt_ly
    sheet['H19'] = goal_13online  # online_goal
    sheet['I19'] = amt_13online  # online_amt
    # HFB13_Index
    sheet['E19'] = index_13_offline[0]
    sheet['F19'] = index_ly_13_offline[0]
    sheet['J19'] = index_13_online
    sheet['K19'] = index_ly_13_online

# HFB14
if len(ly_amt_14offline) > 0 and len(goal_14offline) > 0:
    sheet['B20'] = ly_amt_14offline[0]
    sheet['C20'] = goal_14offline[0]
    sheet['D20'] = amt_14offline[0]
    sheet['G20'] = ly_amt_14online  # online_amt_ly
    sheet['H20'] = goal_14online  # online_goal
    sheet['I20'] = amt_14online  # online_amt
    # HFB14_Index
    sheet['E20'] = index_14_offline[0]
    sheet['F20'] = index_ly_14_offline[0]
    sheet['J20'] = index_14_online
    sheet['K20'] = index_ly_14_online

# HFB15
if len(ly_amt_15offline) > 0 and len(goal_15offline) > 0:
    sheet['B21'] = ly_amt_15offline[0]
    sheet['C21'] = goal_15offline[0]
    sheet['D21'] = amt_15offline[0]
    sheet['G21'] = ly_amt_15online  # online_amt_ly
    sheet['H21'] = goal_15online  # online_goal
    sheet['I21'] = amt_15online  # online_amt
    # HFB15_Index
    sheet['E21'] = index_15_offline[0]
    sheet['F21'] = index_ly_15_offline[0]
    sheet['J21'] = index_15_online
    sheet['K21'] = index_ly_15_online

# HFB16
if len(ly_amt_16offline) > 0 and len(goal_16offline) > 0:
    sheet['B22'] = ly_amt_16offline[0]
    sheet['C22'] = goal_16offline[0]
    sheet['D22'] = amt_16offline[0]
    sheet['G22'] = ly_amt_16online  # online_amt_ly
    sheet['H22'] = goal_16online  # online_goal
    sheet['I22'] = amt_16online  # online_amt
    # HFB16_Index
    sheet['E22'] = index_16_offline[0]
    sheet['F22'] = index_ly_16_offline[0]
    sheet['J22'] = index_16_online
    sheet['K22'] = index_ly_16_online

# HFB17
if len(ly_amt_17offline) > 0 and len(goal_17offline) > 0:
    sheet['B23'] = ly_amt_17offline[0]
    sheet['C23'] = goal_17offline[0]
    sheet['D23'] = amt_17offline[0]
    sheet['G23'] = ly_amt_17online  # online_amt_ly
    sheet['H23'] = goal_17online  # online_goal
    sheet['I23'] = amt_17online  # online_amt
    # HFB17_Index
    sheet['E23'] = index_17_offline[0]
    sheet['F23'] = index_ly_17_offline[0]
    sheet['J23'] = index_17_online
    sheet['K23'] = index_ly_17_online

# HFB18
if len(ly_amt_18offline) > 0 and len(goal_18offline) > 0:
    sheet['B24'] = ly_amt_18offline[0]
    sheet['C24'] = goal_18offline[0]
    sheet['D24'] = amt_18offline[0]
    sheet['G24'] = ly_amt_18online  # online_amt_ly
    sheet['H24'] = goal_18online  # online_goal
    sheet['I24'] = amt_18online  # online_amt
    # HFB18_Index
    sheet['E24'] = index_18_offline[0]
    sheet['F24'] = index_ly_18_offline[0]
    sheet['J24'] = index_18_online
    sheet['K24'] = index_ly_18_online

# HFB70
if len(ly_amt_70offline) > 0 and len(goal_70offline) > 0:
    sheet['B25'] = ly_amt_70offline[0]
    sheet['C25'] = goal_70offline[0]
    sheet['D25'] = amt_70offline[0]
    sheet['G25'] = ly_amt_70online  # online_amt_ly
    sheet['H25'] = goal_70online  # online_goal
    sheet['I25'] = amt_70online  # online_amt
    # HFB18_Index
    sheet['E25'] = index_70_offline[0]
    sheet['F25'] = index_ly_70_offline[0]
    sheet['J25'] = index_70_online
    sheet['K25'] = index_ly_70_online

# 保存更改
workbook.save(file_path2)