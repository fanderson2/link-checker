from bs4 import BeautifulSoup as bs
import datetime
import pandas as pd
import sys
from urllib.request import urlopen
from urllib.error import HTTPError



now = datetime.datetime.now()
xl = pd.ExcelFile("links-to-track.xlsx")
df_client = xl.parse(sys.argv[1])


# Get anchor text
# Multiple links

track_links = {
    "Page URL": [],
    "Link URL": [],
    "Live?":[],
    "Last Check":[]
}


for i, page_url in enumerate(df_client["Page URL"]):
    track_links["Page URL"].append(page_url)
    track_links["Link URL"].append(df_client["Link URL"][i])
    total_rows = i


for i, page_url in enumerate(df_client['Page URL']):
    print(page_url)
    html = urlopen(page_url)
    html_bs = bs(html.read(), "html.parser")

    for a in html_bs.findAll("a", href=True):
        if a['href'] == df_client["Link URL"][i]:
            print("FOUND -->", a, "\t", now.strftime("%d-%m-%Y (%H:%M)"), "\n")
            track_links["Live?"].append("Yes")
            track_links["Last Check"].append(now.strftime("%d-%m-%Y (%H:%M)"))

    for k, v in track_links.items():
        print(len(v))
        if k == "Live?" and len(v) < (i + 1):
            v.append("No")
        if k == "Last Check" and len(v) < (i + 1):
            v.append(now.strftime("%d-%m-%Y (%H:%M)"))




df_final = pd.DataFrame.from_dict(track_links, orient='columns', dtype=None)

writer = pd.ExcelWriter(xl, engine="xlsxwriter")
df_final.to_excel(writer, sheet_name="RT")
writer.save()
