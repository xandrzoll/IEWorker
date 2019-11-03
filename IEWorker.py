import time
from win32com.client import Dispatch


class IEWorker:
    use_auth = True

    def __init__(self, url='', use_auth=True):
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

    def _navigate(self, url):
        self.ie.Navigate(url)
        self.ie_busy()
        self._get_DOM()

    def navigate(self, url):
        self._navigate(url)
        if self.check_auth():
            self.auth()
            self.ie_busy()
        if url != self.auth_url and url != self.ie.LocationURL:
            self._navigate(url)

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

    def set_auth(self,
                 auth_url='',
                 auth_login_id='',
                 auth_pwd_id='',
                 auth_login='',
                 auth_pwd='',
                 auth_button='',
                 ):
        self.use_auth = True
        self.auth_url = auth_url
        self.auth_login_id = auth_login_id
        self.auth_pwd_id = auth_pwd_id
        self.auth_login = auth_login
        self.auth_pwd = auth_pwd
        self.auth_button = auth_button

    def check_auth(self):
        if self.ie.LocationURL == self.auth_url:
            return True
        else:
            return False

    def auth(self):
        login = self.get_elements(id=self.auth_login_id)['elem']
        pwd = self.get_elements(id=self.auth_pwd_id)['elem']
        button = self.get_elements(id=self.auth_button)['elem']
        login.setAttribute('value', self.auth_login)
        pwd.setAttribute('value', self.auth_pwd)
        button.click()

