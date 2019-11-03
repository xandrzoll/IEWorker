import time
from win32com.client import Dispatch


class IEWorker:
    def __init__(self, url=''):
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
        if url:
            self.navigate(url)

    def navigate(self, url):
        self.ie.Navigate(url)
        self.ie_busy()
        self._get_DOM()

    def ie_busy(self):
        while self.ie.Busy:
            time.sleep(0.1)

    def close(self):
        self.ie.Quit()
        self.ie = None

    def _get_DOM(self):
        self._DOM = []
        elems = self.ie.Document.all
        for elem in elems:
            self._DOM.append({
                'elem': elem,
                'tag': elem.tagName,
                'id': elem.getAttribute('id'),
                'name': elem.getAttribute('name'),
            })

    def get_elements(self, id='', name='', cls='', tag=''):
        elems = self._DOM
        elems = list(filter(lambda x: x['tag'] == tag.upper() if tag else x, elems))
        elems = list(filter(lambda x: x['name'] == name if name else x, elems))
        elems = list(filter(lambda x: x['id'] == id if id else x, elems))

        if len(elems) == 1:
            elems = elems[0]
        elif len(elems) == 0:
            elems = None

        return elems
