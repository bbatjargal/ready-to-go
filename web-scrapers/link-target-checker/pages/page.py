from bs4 import BeautifulSoup

from locators.page_locators import PageLocators
from parsers.link import LinkParser

class Page:
    def __init__(self, page):
        self.soup = BeautifulSoup(page, 'html.parser')

    @property
    def FAQ(self):
        locator = PageLocators.FAQLink
        links = self.soup.select(locator)
        return [LinkParser(e) for e in links]