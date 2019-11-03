from IEWorker import *
from config import *


ie = IEWorker()
ie.set_auth(
        auth_url=url_login_form,
        auth_login_id=login_id,
        auth_pwd_id=pwd_id,
        auth_login=iml_login,
        auth_pwd=iml_pwd,
        auth_button='Button1',
    )
ie.navigate('https://office.iml.ru/')

# ie.close()