import time
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC

# Criando variável service 's' que será utilizada no driver para executá-lo
s = Service('msedgedriver.exe')

# Criando variável 'driver' que irá chamar o webdriver Edge
driver = webdriver.Edge(service=s)

# Criando variável 'wait' que irá fazer o webdriver esperar até 20 segundos quando ela for chamada
wait = WebDriverWait(driver, 20)

# Criando variável 'act' para chamar a função ActionChains do webdriver
act = ActionChains(driver)

# Função que aguarda a presença do elemento para clicar nele
def wait_and_click(xpath):
    element = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
    act.move_to_element(element).click().perform()

# Iniciando contagem do tempo de duração da automação
start_time = time.time()

# Maximizando tela do webdriver e entrando no site https://phptravels.com/demo/
driver.maximize_window()
driver.get('https://phptravels.com/demo/')

#Preenchendo formulário 'Demo request form'

# Preenchendo, com o nome informado na variável 'first_name', o campo 'First Name'
first_name = 'Fulano'
driver.find_element(By.XPATH, '//input[@name="first_name"]').send_keys(first_name)
time.sleep(1)
# Preenchendo, com o sobrenome informado na variável 'last_name', o campo 'Last Name'
last_name = 'De Tal'
driver.find_element(By.XPATH, '//input[@name="last_name"]').send_keys(last_name)
time.sleep(1)
# Preenchendo, com o nome informado na variável 'business_name', o campo 'Business Name'
business_name = 'Business Ltda'
driver.find_element(By.XPATH, '//input[@name="business_name"]').send_keys(business_name)
time.sleep(1)
# Preenchendo, com o e-mail informado na variável 'email', o campo 'Email'
email = 'example@email.com'
driver.find_element(By.XPATH, '(//input[@name="email"]) [1]').send_keys(email)
time.sleep(1)
# Realizando a somatória dos números e inserindo o resultado em 'Result ?'
first_number = driver.find_element(By.XPATH, '//span[@id="numb1"]').text
second_number = driver.find_element(By.XPATH, '//span[@id="numb2"]').text
result = int(first_number) + int(second_number)
driver.find_element(By.XPATH, '//input[@id="number"]').send_keys(result)
time.sleep(1)
# Clicando no botão 'Submit'
wait_and_click('//button[@id="demo"]')
time.sleep(2)

# Encerrando contagem de duração da automação
end_time = time.time()

# Atribuindo tempo de duração da automação, em segundos, à variável 'duration_time'
duration_time = int(end_time - start_time)

# Verificando se formulário foi enviado com sucesso ou não
try:
    if driver.find_element(By.XPATH,  '//*[text()=" Thank you!"]').is_displayed():
        # Imprimindo mensagem de preenchimento bem-sucedido e tempo de duração da automação
        print(f"\nForm completed successfully!\n\nAutomation duration: {duration_time} seconds")
except NoSuchElementException:
    # Imprimindo mensagem de falha no preenchimento
    print("Failed to fill out the form.")