from AbreChamados import AbreChamados
import os

import dotenv

dotenv.load_dotenv()

if __name__ == '__main__':

    app = AbreChamados(
        URL         = str(os.getenv('url')),
        USER        = str(os.getenv('username')),
        PASSWORD    = str(os.getenv('password'))
    )


    app._config_chrome_browser()
    app.open_site()
    app.auth()
    app._register_ticket()