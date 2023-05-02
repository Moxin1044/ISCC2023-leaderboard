import re
import requests
from requests.adapters import Retry
from bs4 import BeautifulSoup
from requests.exceptions import Timeout
import pandas as pd
# 看看大家理论题都拿了多少分呢？

print("""
   ___      __     ___     ____  
  |_  )    /  \   |_  )   |__ /  
   / /    | () |   / /     |_ \  
  /___|   _\__/   /___|   |___/  
_|-----|_|-----|_|-----|_|-----| 
"`-0-0-'"`-0-0-'"`-0-0-'"`-0-0-' 
   ___     ___     ___     ___   
  |_ _|   / __|   / __|   / __|  
   | |    \__ \  | (__   | (__   
  |___|   |___/   \___|   \___|  
_|-----|_|-----|_|-----|_|-----| 
"`-0-0-'"`-0-0-'"`-0-0-'"`-0-0-' 
""")
print("【欢迎】：脚本已启动，欢迎使用！ 懒得写多线程了，自求多福！ 当前版本：V1.0")

url_template = "https://iscc.isclab.org.cn/team/{}"
n = 4300
delay = 0.5  # 等待0.5秒后重试

retry_strategy = Retry(
    total=3, # 最多重试3次
    backoff_factor=delay,
    status_forcelist=[500, 502, 503, 504],
    allowed_methods=["HEAD", "GET", "OPTIONS"]
)

rlist = []
for i in range(1, n + 1):
    url = url_template.format(i)
    session = requests.Session()
    session.mount("https://", requests.adapters.HTTPAdapter(max_retries=retry_strategy))
    try:
        response = session.get(url, timeout=2)
        response.raise_for_status()
    except Timeout:
        print("请求超时，等待{}秒后重试".format(delay))
        response = session.get(url, timeout=2)
        response.raise_for_status()
    soup = BeautifulSoup(response.text, 'html.parser')
    h1_tag = soup.find('h1', {'id': 'team-id'})
    if h1_tag is not None:
        team_name = h1_tag.text
        trs = soup.find_all('tr')
        if trs is not None:
            for tr in trs:
                if 'Choice' in tr.text:
                    tds = tr.find_all('td')
                    values = tds[2].text
                    dates = tds[3].text
                else:
                    values = "0"
                    dates = "0"
        else:
            values = "0"
            dates = "0"
        print(f"{team_name} {values} {dates}")
        info = {"ID": i, "名称": team_name, "理论题积分": values, "时间": dates}
        rlist.append(info)

# 按照排名的方式保存为xlsx
df = pd.DataFrame(rlist)
# 将排名列转换为数字类型
df["理论题积分"] = pd.to_numeric(df["理论题积分"])
# 按照排名排序
df = df.sort_values("理论题积分")
# 将 DataFrame 写入 Excel 文件
writer = pd.ExcelWriter("2023-ISCC-理论题排行榜（积分顺序）.xlsx", engine="xlsxwriter")
df.to_excel(writer, index=False, sheet_name="Teams")
writer.close()
print("【注意】：2023-ISCC-理论题排行榜（积分顺序）.xlsx 已生成")

# 按照排名的方式保存为xlsx
df = pd.DataFrame(rlist)
# 将排名列转换为数字类型
df["理论题积分"] = pd.to_numeric(df["理论题积分"])
# 按照排名排序
df = df.sort_values("理论题积分", ascending=False)
# 将 DataFrame 写入 Excel 文件
writer = pd.ExcelWriter("2023-ISCC-理论题排行榜（积分倒序）.xlsx", engine="xlsxwriter")
df.to_excel(writer, index=False, sheet_name="Teams")
writer.close()
print("【注意】：2023-ISCC-理论题排行榜（积分倒序）.xlsx 已生成")
