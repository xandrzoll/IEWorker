from win32com.client import Dispatch


class IEWorker:
    def __init__(self):
        self.ie = Dispatch("InternetExplorer.Application")
        self.ie.Visible = True

