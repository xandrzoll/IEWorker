from win32com.client import Dispatch


class IEWorker:
    def __init__(self, url):
        if url:
            win_id = '{9BA05972-F6A8-11CF-A442-00A0C90A8F39}'
            win_shells = Dispatch(win_id)
            url = url.lower()
            for win_shell in win_shells:
                if url in win_shell.LocationURL.lower():
                    self.ie = win_shell
                    return
        self.ie = Dispatch("InternetExplorer.Application")
        self.ie.Visible = True

    def navigate(self, url):
        self.ie.Navigate(url)

    def get_elements(self, id='', name='', cls='', tag=''):
        pass
