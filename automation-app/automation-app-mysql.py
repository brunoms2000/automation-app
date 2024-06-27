import time
import threading
import pandas as pd
import customtkinter
import mysql.connector
from tkinter import filedialog
from tkinter import messagebox
from selenium import webdriver
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
# Atribuindo variáveis de conexão ao banco de dados
host = "Insira o id do host do banco de dados"
user = "Insira o user do banco de dados"
password_db = "Insira o password do banco de dados"
database = "Insira o nome do banco de dados"
# Atribuindo lista de usuários à variável 'users'
users = ["user1", "user2", "user3", "user4", "user5", "user6"]
# Atribuindo a classe das funções lógicas do sistema (back-end)
class Functions():
    def conecta_bd(self):
        # Para se conectar ao banco de dados
        self.connection = mysql.connector.connect(
            host=host,
            user=user,
            password=password_db,
            database=database,
        )
        self.cursor = self.connection.cursor()
    def desconecta_bd(self):
        # Para se desconectar do banco de dados
        self.cursor.close()
        self.connection.close()
    def create_audit(self, username, response, message):
        # Para criar uma nova informação na tabela audit do banco de dados
        self.conecta_bd()
        self.comando = f'''
        INSERT INTO audit (username, response, message) 
        VALUES ("{username}", "{response}", "{message}")
        '''
        self.cursor.execute(self.comando)
        self.connection.commit()
        self.desconecta_bd()
    def create_users(self, username, password):
        # Para criar uma nova informação nas colunas username e password da tabela users do banco de dados
        self.conecta_bd()
        self.comando = f'''
        INSERT INTO users (username, password)
        VALUES ("{username}", "{password}")
        '''
        self.cursor.execute(self.comando)
        self.connection.commit()
        self.desconecta_bd()
    def create_performance(self, username, automation, clicks, duration_sec, clients_qty):
        # Para criar uma nova informação nas colunas username e password da tabela performance do banco de dados
        self.conecta_bd()
        self.comando = f'''
        INSERT INTO performance (username, automation, clicks, duration_sec, clients_qty)
        VALUES ("{username}", "{automation}", "{clicks}", "{duration_sec}", "{clients_qty}")
        '''
        self.cursor.execute(self.comando)
        self.connection.commit()
        self.desconecta_bd()
    def update_users(self, password, username):
        # Para editar uma informação da tabela users do banco de dados
        self.conecta_bd()
        self.comando = f'''
        UPDATE users
        SET password = "{password}"
        WHERE username = "{username}"
        '''
        self.cursor.execute(self.comando)
        self.connection.commit()
        self.desconecta_bd()
    def read_all_users(self):
        # Para ler a tabela users inteira do banco de dados
        self.conecta_bd()
        self.comando = f'''
        SELECT *
        FROM users
        '''
        self.cursor.execute(self.comando)
        self.df = self.cursor.fetchall()
        self.desconecta_bd()
        return self.df
    def read_username_users(self, username):
        # Para ler a linha de um usuário em específico da tabela users do banco de dados
        self.conecta_bd()
        self.comando = f'''
        SELECT *
        FROM users
        WHERE username = "{username}"
        '''
        self.cursor.execute(self.comando)
        self.df = self.cursor.fetchall()
        self.desconecta_bd()
        return self.df
    def login(self, username_entry, password_entry):
        self.username = username_entry.get()
        self.password = password_entry.get()
        self.users = users
        if self.username in self.users:
            self.users_df = self.read_all_userstest()
            if any(self.username in item for item in self.users_df):
                self.verify = self.read_username_userstest(self.username)
                if any(self.password in item for item in self.verify):
                    self.create_audittests(self.username, "successful login", "Login bem-sucedido")
                    if self.username == "user1":
                        print('open_user1_automations_window')
                    elif self.username == "user2":
                        print('open_user2_automations_window')
                    elif self.username == "user3":
                        print('open_user3_automations_window')
                    elif self.username == "user4":
                        print('open_user4_automations_window')
                    elif self.username == "user5":
                        print('open_user5_automations_window')
                    else:
                        self.automations_test_window()
                else:
                    self.create_audittests(self.username, "login failed",
                                           "Falha no login: senha incorreta")
                    self.login_message_label.configure(text="Senha inválida.", text_color="#D61736")

            else:
                self.create_audittests(self.username, "login failed",
                                       "Falha no login: senha não cadastrada")
                self.login_message_label.configure(
                    text=r'''Senha não cadastrada para o usuário informado, 
cadastrar senha ao clicar em "Registrar Senha".
                    ''', text_color="#D61736")
        else:
            self.create_audittests(self.username, "login failed",
                                   "Falha no login: usuário pré-cadastrado não encontrado")
            self.login_message_label.configure(text='''Usuário pré-cadastrado não encontrado.
            \nPor favor, inserir o usuário pré-cadastrado, 
            designado para seu departamento.
            ''', text_color="#D61736")
    def register_user(self, username_entry, password_entry, confirmation_entry):
        self.username_info = username_entry.get()
        self.password_info = password_entry.get()
        self.confirmation_info = confirmation_entry.get()
        self.users = users
        if self.username_info in self.users:
            self.users_df = self.read_all_userstest()
            if any(self.username_info in item for item in self.users_df):
                self.create_audittests(self.username_info, "register failed",
                                       "Falha no registro: já existe uma senha cadastrada")
                self.register_message_label.configure(text='Já existe uma senha criada para esse usuário.',
                    text_color="#D61736")
            elif self.password_info == self.confirmation_info:
                self.create_userstest(self.username_info, self.password_info)
                self.create_audittests(self.username_info, "successful registration",
                                       "Registro de senha bem-sucedido")
                self.register_message_label.configure(text="Senha salva com sucesso! Pode voltar para a área de login.",
                                             text_color="#002621")
            else:
                self.create_audittests(self.username_info,  "register failed",
                                       "Falha no registro: senha e confirmação de senha incompatíveis")
                self.register_message_label.configure(text="Senha e confirmação de senha não estão compatíveis.",
                    text_color="#D61736")
        else:
            self.create_audittests(self.username_info, "register failed",
                                   "Falha no registro: usuário pré-cadastrado não encontrado")
            self.register_message_label.configure(
                text='''Usuário pré-cadastrado não encontrado.
            \nPor favor, inserir o usuário pré-cadastrado, 
            designado para seu departamento.
            ''', text_color="#D61736")
    def password_redefinition(self, username_entry, password_entry, confirmation_entry):
        self.username_info = username_entry.get()
        self.password_info = password_entry.get()
        self.confirmation_info = confirmation_entry.get()
        self.users = users
        if self.username_info in self.users:
            self.users_df = self.read_all_userstest()
            if any(self.username_info in item for item in self.users_df):
                if self.password_info == self.confirmation_info:
                    self.update_userstest(self.password_info, self.username_info)
                    self.create_audittests(self.username_info, "successful reset",
                                           "Redefinição de senha bem-sucedida")
                    self.redefinition_message_label.configure(
                        text="Senha salva com sucesso! Pode voltar para a área de login.",
                        text_color="#002621"
                    )
                else:
                    self.create_audittests(self.username_info, "reset failed",
                                           "Falha na redefinição: senha e confirmação de senha incompatíveis")
                    self.redefinition_message_label.configure(
                        text='Nova senha e confirmação da nova senha não estão compatíveis.',
                        text_color="#D61736"
                    )
            else:
                self.create_audittests(self.username_info, "reset failed",
                                       "Falha na redefinição: não existe senha cadastrada")
                self.redefinition_message_label.configure(
                    text='''Usuário ainda não possuí senha criada para poder redefini-la, 
cadastrar senha ao clicar em "Registrar Senha" na área de login.''',
                    text_color="#D61736"
                )
        else:
            self.create_audittests(self.username_info, "reset failed",
                                   "Falha na redefinição: usuário pré-cadastrado não encontrado")
            self.redefinition_message_label.configure(
                text='''Usuário pré-cadastrado não encontrado.
            \nPor favor, inserir o usuário pré-cadastrado, 
            designado para seu departamento.''',
                text_color="#D61736"
            )
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
            self.create_audittests(self.username, f"{automation}_interrupt",
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
        # Iniciando contagem do tempo de duração da automação
        self.start_time = time.time()
        self.create_audittests(self.username, "form_filling_start",
                               "Início da execução da automação form filling")
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
                self.create_audittests(self.username, "form_filling_error",
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
        self.create_performancetest(self.username,self.automation,self.total_clicks,self.duration_time, self.total_rows)
        if self.keep_running:
            self.create_audittests(self.username, "form_filling_end",
                                   "Fim da execução da automação form filling")
            self.running_label.configure(text="Preenchimento concluído.")
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
        # Username Title
        usernametitle_label = customtkinter.CTkLabel(window, wraplength=win_width, text="Usuário *",
                                                     text_color="#002621")
        usernametitle_label.pack(padx=10)
        # Username Entry
        username_entry = customtkinter.CTkEntry(window, border_width=1.2, border_color="#002621",
                                                fg_color="#FFFFFF")
        username_entry.pack(padx=10, pady=10)
        # Password Title
        passwordtitle_label = customtkinter.CTkLabel(window, wraplength=win_width, text="Senha *",
                                                     text_color="#002621")
        passwordtitle_label.pack(padx=10)
        # Password Entry
        password_entry = customtkinter.CTkEntry(window, border_width=1.2, border_color="#002621",
                                                fg_color="#FFFFFF",
                                                show="*")
        password_entry.pack(padx=10, pady=10)
        # Forgot Password Button
        forgot_password_button = customtkinter.CTkButton(window, border_width=0, command=self.redefinition_window,
                                                         text="Esqueceu a senha?", text_color="#002621",
                                                         fg_color="#FFFFFF", hover_color="#F1F5F8")
        forgot_password_button.pack(padx=10)
        # Login Button
        login_button = customtkinter.CTkButton(window, border_width=1.2,
                                               text="Login", command=lambda: self.login(username_entry, password_entry),
                                               text_color="#002621", border_color="#002621", fg_color="#FFFFFF",
                                               hover_color="#F1F5F8")
        login_button.pack(padx=10, pady=10)
        # Login Message Label
        self.login_message_label = customtkinter.CTkLabel(window, text="")
        self.login_message_label.pack()
        # Register Button
        register_button = customtkinter.CTkButton(window, border_width=0, command=self.register_window,
                                                  text="Registrar Senha", text_color="#002621",
                                                  fg_color="#FFFFFF", hover_color="#F1F5F8")
        register_button.pack(padx=10, pady=10)
    def register_window(self):
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
            text="Coloque o nome de usuário de seu departamento e depois registre uma senha.", text_color="#002621")
        register_observation_title.pack(padx=10, pady=10)
        # Username Title
        usernameregistertitle_label = customtkinter.CTkLabel(window, wraplength=win_width, text="Usuário",
                                                             text_color="#002621")
        usernameregistertitle_label.pack(padx=10)
        # Username Entry
        usernameregister_entry = customtkinter.CTkEntry(window, border_width=1.2, border_color="#002621",
                                                        fg_color="#FFFFFF")
        usernameregister_entry.pack(padx=10, pady=10)
        # Password Title
        passwordtitleregister_label = customtkinter.CTkLabel(window, wraplength=win_width, text="Senha",
                                                             text_color="#002621")
        passwordtitleregister_label.pack(padx=10)
        # Password Entry
        passwordregister_entry = customtkinter.CTkEntry(window, border_width=1.2, border_color="#002621",
                                                        fg_color="#FFFFFF", show="*")
        passwordregister_entry.pack(padx=10, pady=10)
        # Password Title
        passwordtitleconfirmation_label = customtkinter.CTkLabel(window, wraplength=win_width,
                                                                 text="Confirme a senha",
                                                                 text_color="#002621")
        passwordtitleconfirmation_label.pack(padx=10)
        # Password Confirmation Entry
        passwordconfirmation_entry = customtkinter.CTkEntry(window, border_width=1.2, border_color="#002621",
                                                            fg_color="#FFFFFF", show="*")

        passwordconfirmation_entry.pack(padx=10, pady=10)
        # Save Button
        save_button = customtkinter.CTkButton(window, border_width=1.2, command=lambda: self.register_user(
            usernameregister_entry, passwordregister_entry, passwordconfirmation_entry
        ),
                                              text="Salvar", text_color="#002621", border_color="#002621",
                                              fg_color="#FFFFFF", hover_color="#F1F5F8")
        save_button.pack(padx=10, pady=10)
        # Register Message Label
        self.register_message_label = customtkinter.CTkLabel(window, wraplength=win_width, text="")
        self.register_message_label.pack()
        # Back to Login Window Button
        back_to_login_button = customtkinter.CTkButton(window, border_width=0, command=self.login_window,
                                                       text="Voltar para Login", text_color="#002621",
                                                       border_color="#002621", fg_color="#FFFFFF",
                                                       hover_color="#F1F5F8")
        back_to_login_button.pack(padx=1, pady=1)
    def redefinition_window(self):
        # Garantindo que a janela estará vázia antes de colocar os widgets
        self.clear_window()
        # Configurando a janela
        self.window.geometry('{}x{}'.format(win_width, win_height))
        self.window.title("Automation App")
        self.window.configure(fg_color="#FFFFFF")
        self.window.minsize(win_width, win_height)
        # Redefinition Title
        redefinition_title = customtkinter.CTkLabel(
            window, wraplength=win_width,
            text="Coloque o nome de usuário de seu departamento e depois registre uma nova senha.",
            text_color="#002621"
        )
        redefinition_title.pack(padx=10)
        # Username Title
        redefinition_username_title = customtkinter.CTkLabel(window, wraplength=win_width,
                                                             text="Usuário", text_color="#002621")
        redefinition_username_title.pack(padx=10, pady=10)
        # Username Entry
        redefinition_username_entry = customtkinter.CTkEntry(window, border_width=1.2,
                                                             border_color="#002621",
                                                             fg_color="#FFFFFF")
        redefinition_username_entry.pack(padx=10, pady=10)
        # Password Title
        redefinition_password_title = customtkinter.CTkLabel(window, wraplength=win_width,
                                                             text="Nova senha", text_color="#002621")
        redefinition_password_title.pack(padx=10)
        # Password Entry
        redefinition_password_entry = customtkinter.CTkEntry(window, border_width=1.2,
                                                             border_color="#002621",
                                                             fg_color="#FFFFFF", show="*")
        redefinition_password_entry.pack(padx=10, pady=10)
        # Password Confirmation Title
        redefinition_confirmation_title = customtkinter.CTkLabel(window, wraplength=win_width,
                                                                 text="Confirme a nova senha", text_color="#002621")
        redefinition_confirmation_title.pack(padx=10)
        # Password Confirmation Entry
        redefinition_confirmation_entry = customtkinter.CTkEntry(window, border_width=1.2,
                                                                 border_color="#002621", fg_color="#FFFFFF", show="*")
        redefinition_confirmation_entry.pack(padx=10, pady=10)
        # Save Button
        save_button = customtkinter.CTkButton(window, border_width=1.2, command=lambda: self.password_redefinition(
            redefinition_username_entry, redefinition_password_entry, redefinition_confirmation_entry
        ),
                                              text="Salvar", text_color="#002621", border_color="#002621",
                                              fg_color="#FFFFFF", hover_color="#F1F5F8")
        save_button.pack(padx=10, pady=10)
        # Redefinition Message Label
        self.redefinition_message_label = customtkinter.CTkLabel(window, wraplength=win_width, text="")
        self.redefinition_message_label.pack()
        # Back to Login Window Button
        back_to_login_button = customtkinter.CTkButton(window, border_width=0, command=self.login_window,
                                                       text="Voltar para Login", text_color="#002621",
                                                       border_color="#002621", fg_color="#FFFFFF",
                                                       hover_color="#F1F5F8")
        back_to_login_button.pack(padx=1, pady=1)
    def automations_test_window(self):
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