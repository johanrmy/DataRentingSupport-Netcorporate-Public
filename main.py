import math
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium import webdriver
import pandas as pd
import time
import re
from colorama import Fore, init
from db_management import DbManagement
from vars import *
from execution_time import *
init()
class Renting_data:
    def __init__(self):
        service = Service(executable_path="chromedriver-win64/chromedriver.exe")
        options = Options()
        options.add_experimental_option("detach", True)
        options.page_load_strategy = 'eager'
        self.driver = webdriver.Chrome(service=service, options=options)

    def login(self, username, password):
        try:
            self.driver.get('https://netcorporate.odoo.com/web/login')
            time.sleep(1)
            self.driver.find_element(By.NAME, "login").send_keys(username)
            self.driver.find_element(By.NAME, "password").send_keys(password)
            self.driver.find_elements(By.CLASS_NAME, "btn-block")[0].click()
        except Exception as e:
            self.driver.quit()
            print(Fore.RED + "Error in the login function:", str(e))

    def search_renting(self):
        try:
            self.driver.get(MENU)
            WebDriverWait(self.driver, 3).until(
                EC.presence_of_element_located((By.CLASS_NAME, "o_apps")))
            self.driver.get(RENTING_SUPPORT)
            WebDriverWait(self.driver, 3).until(
                EC.presence_of_element_located((By.CLASS_NAME, "o_content")))

            print(Fore.GREEN + "Renting encontrado")

            total_data = self.driver.find_element(By.XPATH, '//span[@class="o_pager_limit"]').text
            pages = math.ceil(int(total_data) / 80)
            data = []
            data1 = []
            data2 = []
            codigo_list = []
            page = 1
            while pages >= 1:
                table = self.driver.find_element(By.TAG_NAME, "tbody")
                rows = table.find_elements(By.TAG_NAME, "tr")

                for i in range(len(rows)):
                    columns = rows[i].find_elements(By.TAG_NAME, "td")
                    info_no_procesado = columns[1].text
                    info = info_no_procesado.split('/')
                    proceso, serie, pedido_de_alquiler, descripcion = "", "", "", ""
                    for x in info:
                        if x.startswith('S') and not '#' in x:
                            pedido_de_alquiler = x
                        elif any(x.startswith(prefix) for prefix in PROCESO_PREFIX):
                            proceso = x
                        elif x.isalnum() and x.isupper() and len(x) <= 10:
                            serie = x
                        elif '#' in x:
                            descripcion = x
                            codigo = re.search(r'\(#(\d+)\)', x).group(1)
                    aisgnacion = columns[2].find_element(By.TAG_NAME, "span").text
                    cliente_no_procesado = columns[3].find_element(By.TAG_NAME, "a").text
                    if ',' in cliente_no_procesado:
                        cliente, usuario = cliente_no_procesado.strip().split(',', 1)
                    else:
                        cliente = cliente_no_procesado
                        usuario = "Desconocido"
                    etapa = columns[7].text
                    estrellas = columns[5].find_elements(By.CLASS_NAME, "fa-star")
                    estrellas_activas = len(estrellas)
                    if estrellas_activas == 1:
                        prioridad = "Medium priority"
                    elif estrellas_activas == 2:
                        prioridad = "Alta prioridad"
                    elif estrellas_activas == 3:
                        prioridad = "Urgente"
                    else:
                        prioridad = "Desconocida"

                    codigo_list.append(codigo)

                    datos = {'ID': codigo, 'serie': serie, 'proceso': proceso, 'pedido': pedido_de_alquiler,
                            'descripcion': descripcion, 'asignacion': aisgnacion, 'cliente': cliente,
                            'usuario': usuario.strip(),
                            'prioridad': prioridad, 'etapa': etapa}
                    data1.append(datos)
                print(Fore.BLUE + "Página "+ Fore.WHITE + str(page) + Fore.BLUE + " extraído")
                page+=1
                pages -= 1
                next = self.driver.find_element(By.CLASS_NAME, "o_pager_next")
                next.click()
                time.sleep(3)
                WebDriverWait(self.driver, 3).until(
                    EC.staleness_of(table))
            for index, cod in enumerate(codigo_list):
                self.get_ticket(cod)
                time.sleep(2)
                tag_tipo = self.driver.find_element(By.NAME, "ticket_type_id")
                tipo = tag_tipo.find_element(By.TAG_NAME, "span").text
                try:
                    tag_email = self.driver.find_element(By.NAME, "partner_email")
                    email = tag_email.find_element(By.TAG_NAME, "a").text
                except NoSuchElementException:
                    email = "Desconocido"
                try:
                    tag_telefono = self.driver.find_element(By.NAME, "partner_phone")
                    telefono = tag_telefono.find_element(By.TAG_NAME, "a").text
                except NoSuchElementException:
                    telefono = "Desconocido"
                tag_text_ticket = self.driver.find_element(By.CLASS_NAME, "o_readonly")
                descripcion_ticket = tag_text_ticket.get_attribute("innerHTML")
                verificar_cuenta = re.search(r'CUENTA:\s*(&nbsp;)?\s*([^<]+)', descripcion_ticket)
                if verificar_cuenta:
                    if verificar_cuenta.group(1):
                        cuenta = verificar_cuenta.group(2).strip()
                    else:
                        cuenta = verificar_cuenta.group(2).strip()
                else:
                    cuenta = "SIN CUENTA"

                inicio = self.driver.find_elements(By.CLASS_NAME, "o_Message_headerDate")
                fecha_f = inicio[0].get_attribute("title")
                fecha_i = inicio[-1].get_attribute("title")

                datos = {'tipo': tipo, 'email': email, 'telefono': telefono,
                        'descripcion_ticket': descripcion_ticket, 'cuenta': cuenta, 'Creado en': fecha_i,
                        'Ultima Actualizacion': fecha_f}
                data2.append(datos)
                print(Fore.BLUE + "Datos de renting " + Fore.WHITE + cod + Fore.BLUE + " extraído " + Fore.RED + f'[{str(index+1)} de {str(len(codigo_list))}]')

            data = [dict(**d1, **d2) for d1, d2 in zip(data1, data2)]

            dtframe = pd.DataFrame(data)
            name_file = NAME_FILE
            dtframe.to_excel(name_file, index=False)
            print(Fore.GREEN + "Datos renting extraídos exitosamente")

        except Exception as e:
            self.handle_error("Error al ejecutar search_renting:", e)

    def logout(self):
        try:
            user_menu = self.driver.find_element(By.CLASS_NAME, "o_user_menu")
            user_menu.click()

            WebDriverWait(self.driver, 1).until(
                EC.presence_of_element_located((By.CLASS_NAME, "o-dropdown--menu")))

            dropdown_menu = self.driver.find_element(By.CLASS_NAME, "o-dropdown--menu")
            btn_logout = dropdown_menu.find_element(By.CSS_SELECTOR, f'a[href="{LOGOUT}"]')
            btn_logout.click()

            print(Fore.GREEN + "Sesión cerrada")
        except Exception as e:
            print(Fore.RED + "Error al ejecutar logout:", str(e))
        finally:
            self.driver.quit()

    def handle_error(self, msg, exc):
        print(Fore.RED + msg, str(exc))

    def get_ticket(self, id):
        try:
            ticket_url = f"https://netcorporate.odoo.com/web#id={id}&cids=1&menu_id=443&action=615&active_id=2&model=helpdesk.ticket&view_type=form"
            self.driver.get(ticket_url)
            WebDriverWait(self.driver, 3).until(
                EC.presence_of_element_located((By.CLASS_NAME, "o_form_view")))
        except Exception as e:
            self.handle_error("Error al ingresar ticket", e)

if __name__ == "__main__":
    try:
        execute = ExecutionTime()
        start_time = execute.getSomeTime()
        rpa_test = Renting_data()
        rpa_test.login(USERNAME, PASSWORD)
        rpa_test.search_renting()
        end_time = execute.getSomeTime()
        execute_time = execute.setExecutionTime(end_time, start_time)
        execute.getExecutionTime(execute_time)
        time.sleep(1)
        rpa_test.logout()
        data_db = DbManagement(HOST, USER_DB, PASSWORD_DB, DATABASE)
        data_db.connect()
        data_db.save_data_from_excel(NAME_FILE)
        data_db.disconnect()
    except TimeoutException:
        print(Fore.RED + "Error: La página de inicio de sesión no se pudo cargar correctamente.")
