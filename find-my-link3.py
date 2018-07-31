from bs4 import BeautifulSoup as bs
import datetime
import pandas as pd
import sys
from urllib.parse import urlparse
from urllib.request import urlopen
from urllib.error import HTTPError

# Automatically run every month, possibly send email .notifying of dead link

now = datetime.datetime.now()

try:
    xl = pd.ExcelFile("links-to-track.xlsx")
except PermissionError as e:
    print(e)
    sys.exit()


df_client = xl.parse(sys.argv[1])



def getTLD(full_url):
    # Get and return TLD
    parsed_url = urlparse(full_url)
    TLD = parsed_url.scheme + "://" + parsed_url.netloc + "/"
    return TLD


track_links = {
    "Page URL": [],
    "Link URL": [],
    "Anchor":[],
    "Live?":[],
    "Do/No-Follow?":[],
    "Last Check":[]
}


for i, page_url in enumerate(df_client["Page URL"]):
    track_links["Page URL"].append(page_url)
    track_links["Link URL"].append(df_client["Link URL"][i])
    total_rows = i



for i, page_url in enumerate(df_client['Page URL']):
    try:
        html = urlopen(page_url)
        html_bs = bs(html.read(), "html.parser")
    except HTTPError as e:
        print(df_client['Page URL'], "--> ", e)
        track_links["Anchor"].append("-")
        track_links["Do/No-Follow?"].append("-")
        track_links["Live?"].append(str(e))
        track_links["Last Check"].append(now.strftime("%d-%m-%Y (%H:%M)"))

    else:
        for a in html_bs.findAll("a", href=True, rel=True):
            if a['href'] in getTLD(df_client["Link URL"][i]):
                print("FOUND -->", a, "\t", now.strftime("%d-%m-%Y (%H:%M)"), "\n")
                track_links["Anchor"].append(a.contents[0])
                if "nofollow" in a['rel']:
                    track_links["Do/No-Follow?"].append("NoFollow")
                else:
                    track_links["Do/No-Follow?"].append("DoFollow")
                track_links["Live?"].append("Yes")
                track_links["Last Check"].append(now.strftime("%d-%m-%Y (%H:%M)"))

        for k, v in track_links.items():
            if k == "Anchor" and len(v) < (i + 1):
                v.append("-")
            if k == "Do/No-Follow?" and len(v) < (i + 1):
                v.append("-")
            if k == "Live?" and len(v) < (i + 1):
                v.append("No")
            if k == "Last Check" and len(v) < (i + 1):
                v.append(now.strftime("%d-%m-%Y (%H:%M)"))



df_final = pd.DataFrame.from_dict(track_links, orient='columns', dtype=None)

writer = pd.ExcelWriter(xl, engine="xlsxwriter")
df_final.to_excel(writer, sheet_name=sys.argv[1])

try:
    writer.save()
except PermissionError as e:
    print("\n",e)
    print("Please close the document before running this script!")
    sys.exit()
