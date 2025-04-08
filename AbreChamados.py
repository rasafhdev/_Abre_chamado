# type: ignore
from dataclasses import dataclass
import os
import dotenv
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import openpyxl as xl
from time import sleep

dotenv.load_dotenv()

@dataclass
class AbreChamados:
    # não é uma boa pratica pois eu teria que ficar
    # fazendo verificação Not NONE
    # URL      : str | None = os.getenv('url')
    # USER     : str | None = os.getenv('username')
    # PASSWORD : str | None = os.getenv('password')

    URL      : str # url do site
    USER     : str # usuário 
    PASSWORD : str # senha


    def _scraping(self, element_by:By, value:str, action:str, wait:int = 10):

        if action == 'wait':
            #print(f'Raspador aguardando a presença do elemento:{value}')
            return WebDriverWait(self.browser, wait).until(
                EC.presence_of_element_located((element_by, value))
            )
        
        elif action == 'click':
            #print(f'Raspador Tentando clicar no elemento clicavel: {value}')
            element = WebDriverWait(self.browser, wait).until(
                EC.element_to_be_clickable((element_by, value)) 
            )
            element.click()
            return element
        
        elif action == 'find':
            #print(f'Raspador tentanto pesquisando elemento: {value}')
            return self.browser.find_element(element_by, value)
        
    
    def _config_chrome_browser(self) -> webdriver.Chrome:
        ROOT_FOLDER    = Path(__file__).parent
        CHROME_DRIVER  = ROOT_FOLDER / 'drivers' / 'chromedriver'

        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('--start-maximized')

        service = Service(str(CHROME_DRIVER))
        self.driver = webdriver.Chrome(
            service=service,
            options=chrome_options,
        )

        return self.driver


    def open_site(self):
        print('Abrindo Site!')
        self.browser = self._config_chrome_browser()
        self.browser.get(str(self.URL))


    def auth(self):
        print('Realizando autenticação')
        form_input_user = self._scraping(By.NAME, 'UserName', 'wait')
        form_input_pass = self._scraping(By.NAME, 'Password', 'wait')     
        button_submit   = self._scraping(By.ID, 'btnSubmit', 'wait')   
        
        form_input_user.send_keys(self.USER)
        form_input_pass.send_keys(self.PASSWORD)
        button_submit.submit()
        self.browser.get(self.URL)

        print('Sucesso!')
        sleep(2)

        
    def _register_ticket(self):
        sleep(2)
        spreadsheet = xl.load_workbook('chamados_realizados.xlsx')
        page = spreadsheet['Sheet1']

    
        for _, line in enumerate(page.iter_rows(min_row=2)):

            # colunas da planilha
            # Mantive os nomes das colunas em português para manter a igualdade
            (
                solicitante,
                servico,
                categoria,
                urgencia,
                assunto ,
                descricao,
                *_,
            ) = line
            
            

            # caputura dos valores
            _solicitatente = solicitante.value
            _servico       = servico.value
            _categoria     = categoria.value
            _urgencia      = urgencia.value
            _assunto       = assunto.value
            _descricao     = descricao.value


            os.system('clear')
            print(f'Registrando chamado para {_solicitatente}')


            # Solicitante
            person = self._scraping(By.ID, 'select2-chosen-10', 'find')
            person.click()

            search_person = self._scraping(
                By.CSS_SELECTOR,
                'input.select2-focusser',
                'find'
            )
            
            
            search_person.send_keys(_solicitatente)
            sleep(2)
            search_person.send_keys(Keys.ENTER)
            
            
            # Serviço 
            service = self._scraping(
                By.CSS_SELECTOR,
                'a.md-select-treeview-dropdown-field',
                'find')
            service.click()
            sleep(1)

            input_service = self._scraping(
                By.CSS_SELECTOR,
                'input.md-select-treeview-search',
                'find'
            )
            input_service.send_keys(_servico)

            click_on_service = WebDriverWait(self.browser, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//div[text()='9. Outras Solicitações']")
                
                )
            )

            click_on_service.click()

            # jqxWidgetdf35391407e5


            # Categoria
            sleep(1)
            category = self._scraping(By.ID, 'select2-chosen-14', 'find')
            category.click()


            input_category= self._scraping(
                By.CSS_SELECTOR,
                'input.select2-focusser',
                'find'
            )
            sleep(2)
            input_category.send_keys(_categoria)
            sleep(2)
            input_category.send_keys(Keys.DOWN)
            sleep(2)
            input_category.send_keys(Keys.ENTER)

            # Urgência
            sleep(2)
            urgency = self._scraping(By.ID, 'select2-chosen-15', 'find')
            urgency.click()

            select_urgency = self._scraping(
                By.CSS_SELECTOR,
                'input.select2-focusser',
                'find'
            )
            select_urgency.send_keys(_urgencia)
            sleep(2)
            select_urgency.send_keys(Keys.DOWN)
            sleep(2)
            select_urgency.send_keys(Keys.ENTER)



            # Assunto
            form_assunto = self._scraping(By.NAME, 'Subject', 'find')
            form_assunto.click()
            form_assunto.send_keys(_assunto)


            # Descrição
            description_editor = self._scraping(
                By.CLASS_NAME,
                'fr-element',
                'find'
            )
            description_editor.click()
            sleep(1)
            description_editor.send_keys(_descricao)


            # Satus
            button_status = self._scraping(
                By.ID,
                'ticket-status-container',
                'find'
            )
            button_status.click()

            sleep(1)

            status_cursor = self._scraping(
                By.ID,
                's2id_autogen7_search',
                'find'
            )

            status_cursor.click()
            status_cursor.send_keys('Em atendimento')
            status_cursor.send_keys(Keys.ENTER)

            # ATENÇÃO, os scripts estão dentro da env para economizar espaço.
            # clicar no botão de nova ação com o JS
            self.browser.execute_script(os.getenv('nova_acao'))
            sleep(2)

            # clicar no botão salvar com JS
            self.browser.find_element(
                By.CSS_SELECTOR,'button.btn-submit').click()

            print('Save Ticket')

            sleep(3)
            self.browser.get(self.URL)
            sleep(10)


