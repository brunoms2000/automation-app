import time
import random
import requests
import threading
import pandas as pd
import customtkinter
from datetime import datetime
from tkinter import filedialog
from tkinter import messagebox
from selenium import webdriver
import win32com.client as win32
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC
# Window Geometry Format Settings
win_width, win_height = 920, 480
# Atribuindo o Ctk do customtkinter a uma variável window para ser utilizada como padrão das janelas
window = customtkinter.CTk()
# Atribuindo lista de usuários à variável 'users'
users = ["firstname.lastname@email.com",
         "firstname.lastname@email.com",
         "firstname.lastname@email.com",
         "firstname.lastname@email.com",
         "firstname.lastname@email.com"]
group1 = ["firstname.lastname@email.com"]
group2 = ["firstname.lastname@email.com"]
group3 = ["firstname.lastname@email.com"]
group4 = ["firstname.lastname@email.com"]
group5 = ["firstname.lastname@email.com"]
# Atribuindo a classe das funções lógicas do sistema (back-end)
class Functions():
    def create_audit(self, dataehora, username, response, message):
        # Para criar uma nova informação na tabela audit do banco de dados
        try:
            # Cria a estrutura de dados como um dicionário Python
            data_structure = {
                "Datetime": dataehora,
                "User": username,
                "Response": response,
                "Message": message
            }
            # Envia a solicitação POST com os dados JSON
            self.create = requests.post("https://url-do-banco-de-dados-firebase.com/audit/.json",
                                        json=data_structure)
            self.create.raise_for_status()  # Isso levantará uma exceção para códigos de status HTTP de erro
        except requests.exceptions.Timeout:
            print("A requisição excedeu o tempo limite.")
        except requests.exceptions.RequestException as e:
            print(f"Ocorreu um erro na requisição: {e}")
    def create_performance(self, dataehora, username, automation, clicks, duration_sec, clients_qty):
        # Para criar uma nova informação na tabela performance do banco de dados
        try:
            # Cria a estrutura de dados como um dicionário Python
            data_structure = {
                "Datetime": dataehora,
                "User": username,
                "Automation": automation,
                "Clicks": clicks,
                "Duration_sec": duration_sec,
                "Clients_qty": clients_qty
            }
            # Envia a solicitação POST com os dados JSON
            self.create = requests.post("https://url-do-banco-de-dados-firebase.com/performance/.json",
                                   json=data_structure)
            self.create.raise_for_status()  # Isso levantará uma exceção para códigos de status HTTP de erro
        except requests.exceptions.Timeout:
            print("A requisição excedeu o tempo limite.")
        except requests.exceptions.RequestException as e:
            print(f"Ocorreu um erro na requisição: {e}")
    def send_code(self, email):
        # Gerar um número aleatório de seis dígitos
        self.numero_aleatorio = random.randint(100000, 999999)
        self.recipient = email
        # Criando integração com o outlook
        self.outlook = win32.Dispatch('outlook.application')
        # Criando e-mail
        self.email = self.outlook.CreateItem(0)
        # Configurando as informações do seu e-mail
        self.email.To = self.recipient
        self.email.Subject = 'Automation App - Código de confirmação'
        self.email.HTMLBody = f'''
        <p>Olá,</p>
        <p>Seu código de confirmação é: <strong>{self.numero_aleatorio}</strong></p>
        <p><strong>Não compartilhe esse cógido com mais ninguém.</strong></p>
        <p>Att,</p>
        <p>Automation App.</p>
        '''
        self.email.Send()
        # Extraímos o nome e o sobrenome dividindo a string primeiro pelo '@' e depois pelo '.'
        name_parts = self.useremail.split('@')[0].split('.')
        # Capitalizamos a primeira letra de cada parte
        capitalized_parts = [part.capitalize() for part in name_parts]
        # Juntamos as partes com um espaço
        self.user = ' '.join(capitalized_parts)
        # Obtém a data e hora atuais
        now = datetime.now()
        # Formata a data e a hora
        self.dataehora = now.strftime("%Y-%m-%d %H:%M:%S")
        self.create_audit(self.dataehora, self.user, "code sent", "Código de confirmação enviado")
        self.confirmation_window()
    def code_verification(self, code):
        self.group1 = ["firstname.lastname@email.com"]
        self.group2 = ["firstname.lastname@email.com"]
        self.group3 = ["firstname.lastname@email.com"]
        self.group4 = ["firstname.lastname@email.com"]
        self.group5 = ["firstname.lastname@email.com"]
        self.code_user = code.get()
        # Extraímos o nome e o sobrenome dividindo a string primeiro pelo '@' e depois pelo '.'
        name_parts = self.useremail.split('@')[0].split('.')
        # Capitalizamos a primeira letra de cada parte
        capitalized_parts = [part.capitalize() for part in name_parts]
        # Juntamos as partes com um espaço
        self.user = ' '.join(capitalized_parts)
        if str(self.code_user) == str(self.numero_aleatorio):
            # Obtém a data e hora atuais
            now = datetime.now()
            # Formata a data e a hora
            self.dataehora = now.strftime("%Y-%m-%d %H:%M:%S")
            self.create_audit(self.dataehora, self.user, "successful login", "Login bem-sucedido")
            if self.useremail in self.group1:
                print("self.window1()")
            elif self.useremail in self.group2:
                print("self.window2()")
            elif self.useremail in self.group3:
                print("self.window3()")
            elif self.useremail in self.group4:
                print("self.window4()")
            elif self.useremail in self.group5:
                print("self.window5()")
            else:
                self.automations_test_window()
        else:
            # Obtém a data e hora atuais
            now = datetime.now()
            # Formata a data e a hora
            self.dataehora = now.strftime("%Y-%m-%d %H:%M:%S")
            self.create_audit(self.dataehora, self.user, "login failed",
                              "Falha no login: código incorreto")
            self.confirmation_message_label.configure(text='Número inválido.', text_color="#D61736")
    def login(self, useremail_entry):
        self.useremail = useremail_entry.get()
        self.users = users
        if self.useremail.endswith("@email.com"):
            if self.useremail in self.users:
                self.send_code(self.useremail)
            else:
                self.login_message_label.configure(text='''Acesso negado.\n
                Verifique se o e-mail foi digitado corretamente. Caso o e-mail esteja correto, contatar equipe responsável.
                ''', text_color="#D61736")
        else:
            self.login_message_label.configure(text="E-mail ou texto inserido não condiz com um e-mail válido.",
                                               text_color="#D61736")
    def upload_file(self):
        self.file = filedialog.askopenfilename(title="Selecione o arquivo",
                                               filetypes=[("Arquivo excel", "*.xlsx")])
        if self.file:
            # Usuário selecionou um arquivo
            self.finishupload_label.configure(text="Upload concluído.")
        else:
            #Usuário não selecionou um arquivo
            self.finishupload_label.configure(text="")
    def runcode_in_thread(self, runcode):
        self.status = messagebox.askyesno(
            title="Atenção",
            message="Você realmente deseja executar a automação?"
        )
        if self.status == True:
            self.thread = threading.Thread(target=runcode)
            self.thread.start()
    def stop_run(self, automation):
        self.status = messagebox.askyesno(
            title="Atenção",
            message="Você realmente deseja parar a automação?"
        )
        if self.status == True:
            self.keep_running = False
            # Extraímos o nome e o sobrenome dividindo a string primeiro pelo '@' e depois pelo '.'
            name_parts = self.useremail.split('@')[0].split('.')
            # Capitalizamos a primeira letra de cada parte
            capitalized_parts = [part.capitalize() for part in name_parts]
            # Juntamos as partes com um espaço
            self.user = ' '.join(capitalized_parts)
            # Obtém a data e hora atuais
            now = datetime.now()
            # Formata a data e a hora
            self.dataehora = now.strftime("%Y-%m-%d %H:%M:%S")
            self.create_audit(self.dataehora, self.user, f"{automation}_interrupt",
                                   f"Execução da automação {automation} interrompida pelo usuário")
            self.stoprunning_label.configure(
                text="Automação interrompida conforme solicitado.",
                text_color="#D61736"
            )
    def move_and_click(self, xpath):
        self.element = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
        ActionChains(self.driver).move_to_element(self.element).click().perform()
    def runcode_form_filling(self):
        # Ler base em excel
        self.df = pd.read_excel(self.file, dtype=str)
        # Atribuindo configurações do webdriver às variáveis
        self.s = Service('msedgedriver.exe')
        self.driver = webdriver.Edge(service=self.s)
        self.wait = WebDriverWait(self.driver, 20)
        # Variável de controle para interromper a função runcode
        self.keep_running = True
        self.automation = 'Form_Filling'
        # Extraímos o nome e o sobrenome dividindo a string primeiro pelo '@' e depois pelo '.'
        name_parts = self.useremail.split('@')[0].split('.')
        # Capitalizamos a primeira letra de cada parte
        capitalized_parts = [part.capitalize() for part in name_parts]
        # Juntamos as partes com um espaço
        self.user = ' '.join(capitalized_parts)
        # Obtém a data e hora atuais
        now = datetime.now()
        # Formata a data e a hora
        self.dataehora = now.strftime("%Y-%m-%d %H:%M:%S")
        self.create_audit(self.dataehora, self.user, "form_filling_start",
                               "Início da execução da automação form filling")
        # Iniciando contagem do tempo de duração da automação
        self.start_time = time.time()
        self.running_label.configure(text="Acessando site...")
        self.percentage.configure(text="0%")
        # Maximizando tela do webdriver e entrando no site https://phptravels.com/demo/
        self.driver.maximize_window()
        self.driver.get('https://phptravels.com/demo/')
        time.sleep(1)
        # Contabilizando número de cliques realizados
        self.total_clicks = 0
        self.total_rows = len(self.df)
        for index, row in self.df.iterrows():
            try:
                new_index = str(int(index + 1))
                new_total_rows = str(int(self.total_rows))
                self.running_label.configure(
                    text="Linha " + new_index + " de " + new_total_rows + ".\n\nCliente: " + row["Nome"])
                if not self.keep_running:
                    break
                # Clicando em 'First Name' de 'Instant demo request form' e
                # preenchendo com o nome informado na variável 'first_name'
                first_name = row["Nome"]
                self.move_and_click('//input[@name="first_name"]')
                self.total_clicks = self.total_clicks + 1
                time.sleep(1)
                if not self.keep_running:
                    break
                self.driver.find_element(By.XPATH, '//input[@name="first_name"]').send_keys(first_name)
                time.sleep(1)
                if not self.keep_running:
                    break
                # Clicando em 'Last Name' de 'Instant demo request form' e
                # preenchendo com o nome informado na variável 'last_name'
                last_name = row["Sobrenome"]
                self.move_and_click('//input[@name="last_name"]')
                self.total_clicks = self.total_clicks + 1
                time.sleep(1)
                if not self.keep_running:
                    break
                self.driver.find_element(By.XPATH, '//input[@name="last_name"]').send_keys(last_name)
                time.sleep(1)
                if not self.keep_running:
                    break
                # Clicando em 'Business Name' de 'Instant demo request form' e
                # preenchendo com o nome informado na variável 'business_name'
                business_name = row["Empresa"]
                self.move_and_click('//input[@name="business_name"]')
                self.total_clicks = self.total_clicks + 1
                time.sleep(1)
                if not self.keep_running:
                    break
                self.driver.find_element(By.XPATH, '//input[@name="business_name"]').send_keys(business_name)
                time.sleep(1)
                if not self.keep_running:
                    break
                # Clicando em 'Email' de 'Instant demo request form' e
                # preenchendo com o nome informado na variável 'email'
                email = row["Email"]
                self.move_and_click('(//input[@name="email"]) [1]')
                self.total_clicks = self.total_clicks + 1
                time.sleep(1)
                if not self.keep_running:
                    break
                self.driver.find_element(By.XPATH, '(//input[@name="email"]) [1]').send_keys(email)
                time.sleep(1)
                if not self.keep_running:
                    break
                # Realizando a somatória dos números e inserindo o resultado em 'Result ?'
                first_number = self.driver.find_element(By.XPATH, '//span[@id="numb1"]').text
                second_number = self.driver.find_element(By.XPATH, '//span[@id="numb2"]').text
                result = int(first_number) + int(second_number)
                self.driver.find_element(By.XPATH, '//input[@id="number"]').send_keys(result)
                time.sleep(1)
                # Clicando no botão 'Submit'
                self.move_and_click('//button[@id="demo"]')
                time.sleep(5)
                # Verificando se formulário foi enviado com sucesso ou não
                try:
                    if self.driver.find_element(By.XPATH, '//*[text()=" Thank you!"]').is_displayed():
                        # Imprimindo mensagem de preenchimento bem-sucedido e tempo de duração da automação
                        print(f"\nForm completed successfully!")
                except NoSuchElementException:
                    # Imprimindo mensagem de falha no preenchimento
                    print("Failed to fill out the form.")
                # Calcular o percentual de progresso
                progress_percent = (index + 1) / self.total_rows * 100
                # Atualizar o valor do percentual de progresso
                per = str(int(progress_percent))
                self.percentage.configure(text=per + "%")
                self.percentage.update()
                # Atualizar o valor da barra de progresso
                self.progress_bar.set(float(progress_percent) / 100)
            except:
                # Obtém a data e hora atuais
                now = datetime.now()
                # Formata a data e a hora
                self.dataehora = now.strftime("%Y-%m-%d %H:%M:%S")
                self.create_audit(self.dataehora, self.user, "form_filling_error",
                                       "Falha na execução da automação form filling")
                self.keep_running = False
                self.stoprunning_label.configure(
                    text="Falha na execução da automação.",
                    text_color="#D61736"
                )
                break
        self.driver.quit()
        # Encerrando contagem de duração do código
        self.end_time = time.time()
        # Atribuindo tempo de duração em segundos do código à variável 'duration_time'
        self.duration_time = int(self.end_time - self.start_time)
        # Obtém a data e hora atuais
        now = datetime.now()
        # Formata a data e a hora
        self.dataehora = now.strftime("%Y-%m-%d %H:%M:%S")
        self.create_performance(self.dataehora,self.user,self.automation,self.total_clicks,self.duration_time, self.total_rows)
        if self.keep_running:
            self.running_label.configure(text="Preenchimento concluído.")
            # Obtém a data e hora atuais
            now = datetime.now()
            # Formata a data e a hora
            self.dataehora = now.strftime("%Y-%m-%d %H:%M:%S")
            self.create_audit(self.dataehora, self.user, "form_filling_end",
                                   "Fim da execução da automação form filling")
# Atribuindo a classe de interface de usuário do sistema (front-end)
class User_Interface(Functions):
    def __init__(self):
        self.window = window
        self.login_window()
        self.loop()
    def loop(self):
        # Inicia o loop da interface gráfica
        self.window.mainloop()
    def login_window(self):
        # Garantindo que a janela estará vázia antes de colocar os widgets
        self.clear_window()
        # Configurando a janela
        self.window.geometry('{}x{}'.format(win_width, win_height))
        self.window.title("Automation App")
        self.window.configure(fg_color="#FFFFFF")
        self.window.minsize(win_width, win_height)
        # Title
        title_label = customtkinter.CTkLabel(window, wraplength=win_width, text="Login", text_color="#002621")
        title_label.pack(padx=10, pady=10)
        # User e-mail Title
        useremailtitle_label = customtkinter.CTkLabel(window, wraplength=win_width, text="E-mail",
                                                      text_color="#002621")
        useremailtitle_label.pack(padx=10)
        # User e-mail Entry
        useremail_entry = customtkinter.CTkEntry(window, border_width=1.2, border_color="#002621",
                                                 fg_color="#FFFFFF")
        useremail_entry.pack(padx=10, pady=10)
        # Login Button
        login_button = customtkinter.CTkButton(window, border_width=1.2,
                                               text="Login", command=lambda: self.login(useremail_entry),
                                               text_color="#002621", border_color="#002621", fg_color="#FFFFFF",
                                               hover_color="#F1F5F8")
        login_button.pack(padx=10, pady=10)
        # Login Message Label
        self.login_message_label = customtkinter.CTkLabel(window, text="")
        self.login_message_label.pack()
    def confirmation_window(self):
        # Garantindo que a janela estará vázia antes de colocar os widgets
        self.clear_window()
        # Configurando a janela
        self.window.geometry('{}x{}'.format(win_width, win_height))
        self.window.title("Automation App")
        self.window.configure(fg_color="#FFFFFF")
        self.window.minsize(win_width, win_height)
        # Regiser title
        register_observation_title = customtkinter.CTkLabel(
            window, wraplength=win_width,
            text="Informe o código de confirmação no campo abaixo.", text_color="#002621")
        register_observation_title.pack(padx=10, pady=10)
        # Confirmation Title
        confirmationtitle_label = customtkinter.CTkLabel(window, wraplength=win_width,
                                                         text="Um código de confirmação foi enviado para seu e-mail. Por favor, inserir no campo abaixo.",
                                                         text_color="#002621")
        confirmationtitle_label.pack(padx=10)
        # Confirmation Entry
        confirmation_entry = customtkinter.CTkEntry(window, border_width=1.2, border_color="#002621",
                                                    fg_color="#FFFFFF")
        confirmation_entry.pack(padx=10, pady=10)
        # Confirmation Message Label
        self.confirmation_message_label = customtkinter.CTkLabel(window, wraplength=win_width, text="")
        self.confirmation_message_label.pack()
        # Next Button
        next_button = customtkinter.CTkButton(window, border_width=1.2,
                                              command=lambda: self.code_verification(confirmation_entry),
                                              text="Avançar", text_color="#002621", border_color="#002621",
                                              fg_color="#FFFFFF", hover_color="#F1F5F8")
        next_button.pack(padx=10, pady=10)
        # Resend Button
        resend_button = customtkinter.CTkButton(window, border_width=0, command=lambda: self.send_code(self.useremail),
                                                text="Reenviar código", text_color="#002621",
                                                fg_color="#FFFFFF", hover_color="#F1F5F8")
        resend_button.pack(padx=1, pady=1)
    def automations_test_window(self):
        # Obtém a data e hora atuais
        now = datetime.now()
        # Formata a data e a hora
        self.dataehora = now.strftime("%Y-%m-%d %H:%M:%S")
        self.create_audit(self.dataehora, self.user, "window opened",
                          "Acesso à janela Automations Test bem-sucedido")
        # Garantindo que a janela estará vázia antes de colocar os widgets
        self.clear_window()
        # Configurando a janela
        self.window.geometry('{}x{}'.format(win_width, win_height))
        self.window.title("Automation App - Test")
        self.window.configure(fg_color="#FFFFFF")
        self.window.minsize(win_width, win_height)
        # Title
        title_label = customtkinter.CTkLabel(window, wraplength=win_width,
                                             text="Escolha abaixo a automação que deseja acessar", text_color="#002621")
        title_label.pack(padx=10, pady=10)
        # Form Filling Button
        form_filling_button = customtkinter.CTkButton(window, border_width=2,
                                                      command=self.form_filling_window,
                                                      text="Form Filling",
                                                      text_color="#002621", border_color="#002621", fg_color="#FFFFFF",
                                                      hover_color="#F1F5F8")
        form_filling_button.pack(padx=10, pady=10)
    def form_filling_window(self):
        # Obtém a data e hora atuais
        now = datetime.now()
        # Formata a data e a hora
        self.dataehora = now.strftime("%Y-%m-%d %H:%M:%S")
        self.create_audit(self.dataehora, self.user, "window opened",
                          "Acesso à janela Form Filling window bem-sucedido")
        # Garantindo que a janela estará vázia antes de colocar os widgets
        self.clear_window()
        # Configurando a janela
        self.window.geometry('{}x{}'.format(win_width, win_height))
        self.window.title("Automation App - Preenchimento de formulário teste")
        self.window.configure(fg_color="#FFFFFF")
        self.window.minsize(win_width, win_height)
        # Title
        title_Label = customtkinter.CTkLabel(window, wraplength=win_width,
                                             text="Faça abaixo o upload da base teste",
                                             text_color="#002621")
        title_Label.pack(padx=10, pady=10)
        # Insert File
        insertfile_button = customtkinter.CTkButton(window, border_width=2, command=self.upload_file,
                                                    text="Upload", text_color="#002621",
                                                    border_color="#002621", fg_color="#FFFFFF", hover_color="#F1F5F8")
        insertfile_button.pack(padx=10, pady=10)
        # Finished Upload
        self.finishupload_label = customtkinter.CTkLabel(window, text="", text_color="#002621")
        self.finishupload_label.pack()
        # Run Button
        run_button = customtkinter.CTkButton(window, border_width=2,
                                             command=lambda: self.runcode_in_thread(self.runcode_form_filling),
                                             text="Executar", text_color="#002621",
                                             border_color="#002621", fg_color="#FFFFFF", hover_color="#F1F5F8")
        run_button.pack(padx=10, pady=10)
        # Running label
        self.running_label = customtkinter.CTkLabel(window, text="", text_color="#002621")
        self.running_label.pack()
        # Stop Run Button
        stop_button = customtkinter.CTkButton(window, text="Parar", command=lambda: self.stop_run('form_filling'),
                                              fg_color="#D61736", hover_color="#AE122C")
        stop_button.pack(padx=10, pady=10)
        # ProgressBar and percentage
        self.progress_bar = customtkinter.CTkProgressBar(window, width=400, progress_color="#85DDA4")
        self.progress_bar.set(0)
        self.progress_bar.pack(padx=10, pady=10)
        self.percentage = customtkinter.CTkLabel(window, text="", text_color="#002621")
        self.percentage.pack()
        # Stop Running label
        self.stoprunning_label = customtkinter.CTkLabel(window, text="")
        self.stoprunning_label.pack()
        # Back to Automations Window Button
        back_to_automations_button = customtkinter.CTkButton(window, border_width=0,
                                                             command=self.automations_test_window,
                                                             text="Voltar", text_color="#002621",
                                                             border_color="#002621", fg_color="#FFFFFF",
                                                             hover_color="#F1F5F8")
        back_to_automations_button.pack(padx=1, pady=1)
    def clear_window(self):
        for widget in self.window.winfo_children():
            widget.destroy()
User_Interface()