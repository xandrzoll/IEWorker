from win32com.client import Dispatch


class IEWorker:
    def __init__(self):
        self.ie = Dispatch("InternetExplorer.Application")
        self.ie.Visible = True

    def navigate(self, url):
        self.ie.Navigate(url)

    def get_elements(self, id='', name='', cls='', tag=''):
        pass
