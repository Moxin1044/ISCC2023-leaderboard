import requests
import re
from bs4 import BeautifulSoup
import pandas as pd
# 单独跑中小学生赛区
# NYCC人不多，所以慢

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


# 设置要访问的URL和循环次数
# url_template = "https://iscc.isclab.org.cn/team/{}" # 练武题接口
# url_template = "https://iscc.isclab.org.cn/teamarena/{}" # 擂台赛接口
url_template = "https://iscc.isclab.org.cn/team/{}"
n = 4300

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
        if team_name[:5] == "nycc-":
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
writer = pd.ExcelWriter("2023-ISCC-中小学赛区-练武题排行榜（排名顺序）.xlsx", engine="xlsxwriter")
df.to_excel(writer, index=False, sheet_name="Teams")
writer.close()
print("【注意】：2023-ISCC--中小学赛区-练武题排行榜（排名顺序）.xlsx 已生成")
# 按照ID的方式保存为xlsx
df = pd.DataFrame(rlist)
# 将排名列转换为数字类型
df["ID"] = pd.to_numeric(df["ID"])
# 按照排名排序
df = df.sort_values("ID")
# 将 DataFrame 写入 Excel 文件
writer = pd.ExcelWriter("2023-ISCC-中小学赛区-练武题排行榜（ID顺序）.xlsx", engine="xlsxwriter")
df.to_excel(writer, index=False, sheet_name="Teams")
writer.close()
print("【注意】：2023-ISCC-中小学赛区-练武题排行榜（ID顺序）.xlsx 已生成")
print("【注意】：练武排行榜已生成，正在生成擂台赛排行榜！")

leitai_url_template = "https://iscc.isclab.org.cn/teamarena/{}" # 擂台赛接口
# 循环访问URL并检查是否存在指定的<h1>标签
r1list = []
for i in range(1, n + 1):
    url = leitai_url_template.format(i)
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')
    h1_tag = soup.find('h1', {'id': 'team-id'})
    h3_tag = soup.find('h3', {'class': 'text-center'})
    if h1_tag is not None:
        team_name = h1_tag.text
        if team_name[:5] == "nycc-":
            h3_tag_text = h3_tag.text
            pattern = r'总积分为:(\d+),排在(\d+)位。'
            match = re.search(pattern, h3_tag_text)
            if match:
                total_points = int(match.group(1))
                rank = int(match.group(2))
                info = {"ID": i, "名称": team_name, "积分": total_points, "排名": rank}
                print(info)
                r1list.append(info)
            else:
                info = {"ID": i, "名称": team_name, "积分": "0", "排名": "0"}
                print(info)
                r1list.append(info)


# 按照排名的方式保存为xlsx
df = pd.DataFrame(r1list)
# 将排名列转换为数字类型
df["排名"] = pd.to_numeric(df["排名"])
# 按照排名排序
df = df.sort_values("排名")
# 将 DataFrame 写入 Excel 文件
writer = pd.ExcelWriter("2023-ISCC-中小学赛区-擂台赛排行榜（排名顺序）.xlsx", engine="xlsxwriter")
df.to_excel(writer, index=False, sheet_name="Teams")
writer.close()
print("【注意】：2023-ISCC-中小学赛区-擂台赛排行榜（排名顺序）.xlsx 已生成")
# 按照ID的方式保存为xlsx
df = pd.DataFrame(r1list)
# 将排名列转换为数字类型
df["ID"] = pd.to_numeric(df["ID"])
# 按照排名排序
df = df.sort_values("ID")
# 将 DataFrame 写入 Excel 文件
writer = pd.ExcelWriter("2023-ISCC-中小学赛区-擂台赛排行榜（ID顺序）.xlsx", engine="xlsxwriter")
df.to_excel(writer, index=False, sheet_name="Teams")
writer.close()
print("【注意】：2023-ISCC-中小学赛区-擂台赛排行榜（ID顺序）.xlsx 已生成")
print("【注意】：已全部完成！")