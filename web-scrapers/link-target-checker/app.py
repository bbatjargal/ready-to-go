import requests

from pages.page import Page

page_content = requests.get("--HERE IS YOUR LINK--").content
page = Page(page_content)

import csv

with open("Hyperlinks-on-page.csv", "w") as f:
    f.write('Text,target,Need to be fixed,href\n')
    for link in page.FAQ:
        f.write(link.csvline)
