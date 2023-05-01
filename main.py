import re
import requests
from bs4 import BeautifulSoup
import pandas as pd

# 设置要访问的URL和循环次数
# url_template = "https://iscc.isclab.org.cn/team/{}" # 练武题接口
# url_template = "https://iscc.isclab.org.cn/teamarena/{}" # 擂台赛接口
url_template = "https://iscc.isclab.org.cn/team/{}"
n = 50

# 循环访问URL并检查是否存在指定的<h1>标签
rlist = []
for i in range(1, n + 1):
    url = url_template.format(i)
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')
    h1_tag = soup.find('h1', {'id': 'team-id'})
    h3_tag = soup.find('h3', {'class': 'text-center'})
    if h1_tag is not None:
        team_name = h1_tag.text
        h3_tag_text = h3_tag.text
        pattern = r'总积分为:(\d+),排在(\d+)位。'
        match = re.search(pattern, h3_tag_text)
        if match:
            total_points = int(match.group(1))
            rank = int(match.group(2))
            info = {"ID": i, "名称": team_name, "积分": total_points, "排名": rank}
            print(info)
            rlist.append(info)
        else:
            info = {"ID": i, "名称": team_name, "积分": "0", "排名": "0"}
            print(info)
            rlist.append(info)

# 按照排名的方式保存为xlsx
df = pd.DataFrame(rlist)
# 将排名列转换为数字类型
df["排名"] = pd.to_numeric(df["排名"])
# 按照排名排序
df = df.sort_values("排名")
# 将 DataFrame 写入 Excel 文件
writer = pd.ExcelWriter("2023-ISCC-练武题排行榜（排名顺序）.xlsx", engine="xlsxwriter")
df.to_excel(writer, index=False, sheet_name="Teams")
writer.close()

# 按照ID的方式保存为xlsx
df = pd.DataFrame(rlist)
# 将排名列转换为数字类型
df["ID"] = pd.to_numeric(df["ID"])
# 按照排名排序
df = df.sort_values("ID")
# 将 DataFrame 写入 Excel 文件
writer = pd.ExcelWriter("2023-ISCC-练武题排行榜（ID顺序）.xlsx", engine="xlsxwriter")
df.to_excel(writer, index=False, sheet_name="Teams")
writer.close()
