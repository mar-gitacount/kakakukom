import requests
from bs4 import BeautifulSoup

# スクレイピング対象のURL
# url = "https://kakaku.com/watch_accessory/watch/itemlist.aspx?pdf_ma=5090&pdf_Spec106=1,2&pdf_vi=c"
url = "https://kakaku.com/item/K0001529700/"
# url = "https://c.kakaku.com/forwarder/forward.aspx?ShopCD=7279&PrdKey=K0001529700&Url=https%3A%2F%2Fhousekihiroba%2Ejp%2Fshop%2Fg%2FgRX3150&Hash=0445cbd7b11e1839ce159fbc0cb7d3f4"

# requestsを使ってHTMLを取得
response = requests.get(url)

# ステータスコードが200（成功）でない場合はエラーを表示
if response.status_code != 200:
    print(f"エラー： {response.status_code}")
    exit()

# BeautifulSoupを使ってHTMLを解析
soup = BeautifulSoup(response.text, "html.parser")


# class="ckitemLink"の要素を取得
ckitem_links = soup.find_all(class_="ckitemLink")

# <tbody> タグを取得
tbody_tag = soup.find("tbody")

# <tbody> タグ内のテキストを抽出して表示
if tbody_tag:
    tbody_text = tbody_tag.get_text(strip=True)
    print(tbody_text)
else:
    print("<tbody> タグが見つかりませんでした。")


# 取得した各要素からテキストを抽出して表示
for link in ckitem_links:
    text_content = link.get_text(strip=True)
    print(text_content)

# # 取得した要素を表示
# for link in ckitem_links:
#     print(link.get("href"))


# # 解析したHTMLから情報を取得
# # 例えば、すべてのリンクを取得する場合:
# for link in soup.find_all("a"):
#     print(link.get("href"))
