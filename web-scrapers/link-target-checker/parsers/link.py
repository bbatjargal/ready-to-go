from locators.links_locators import LinkLocators


class LinkParser:
    def __init__(self, parent):
        self.parent = parent

    def __repr__(self):
        return f'{{ Link: "{self.text}", target: "{self.target}", href: "{self.href}" }}'


    @property
    def csvline(self):
        appropraite_target = ''
        if self.target != '_blank' and (self.href.startswith('http://') or self.href.startswith('https://')):
            appropraite_target = '_blank'
        return f'"{self.text}","{self.target}","{appropraite_target}","{self.href}"\n'

    @property
    def href(self):
        return self.parent.attrs['href']

    @property
    def target(self):
        if 'target' in self.parent.attrs:
            return self.parent.attrs['target']
        return '--NOT DEFINED--'

    @property
    def text(self):
        locator = LinkLocators.Text
        el = self.parent.select_one(locator)
        if el:
            return el.string
        return self.parent.string

    @property
    def links(self):
        locator = LinkLocators.Text
        return [e.string for e in self.parent.select(locator)]