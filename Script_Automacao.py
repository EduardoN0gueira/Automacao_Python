from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
import openpyxl

#iniciar o selenium com o comando webdriver.Chrome
#e colocamos em uma variável
driver = webdriver.Chrome()
#colocamos o link do site

driver.get('https://contabilidade-devaprender.netlify.app/')
#pausa de 5 segundos para carregar o site
sleep(3)

#xpath, oque é? o xpath é o caminho
# tag[@atributo='valor']

#após procurar o caminho, colocamos o XPATH
email = driver.find_element(By.XPATH,"//input[@id='email']")

#detalhe: use aspas diferentes!!

sleep(2)
email.send_keys('admin@contabilidade.com')

# agora repetimos o processo para a senha

senha = driver.find_element(By.XPATH,"//input[@id='senha']")
sleep(2)
senha.send_keys('teste123')

# agora precisamos apertar no botão de entrar

botao_entrar = driver.find_element(By.XPATH,"//button[@id='Entrar']")

botao_entrar.click()

sleep(4)

# vamos agora, retirar as informçôes da planilha

empresas = openpyxl.load_workbook('./empresas.xlsx')
pagina_empresas = empresas['dados empresas']

# vamos ler linha por linha, a partir da linha 2
for linha in pagina_empresas.iter_rows(min_row=2,values_only=True):
    nome_empresa, email, telefone, endereco, cnpj, area_atuacao, quantidade_de_funcionarios, data_fundacao = linha

    driver.find_element(By.ID,'nomeEmpresa').send_keys(nome_empresa)
    sleep(2)
    driver.find_element(By.ID,'emailEmpresa').send_keys(email)
    sleep(2)
    driver.find_element(By.ID,'telefoneEmpresa').send_keys(telefone)
    sleep(2)
    driver.find_element(By.ID,'enderecoEmpresa').send_keys(endereco)
    sleep(2)
    driver.find_element(By.ID,'cnpj').send_keys(cnpj)
    sleep(2)
    driver.find_element(By.ID,'areaAtuacao').send_keys(area_atuacao)
    sleep(2)
    driver.find_element(By.ID,'numeroFuncionarios').send_keys(quantidade_de_funcionarios)
    sleep(2)
    driver.find_element(By.ID,'dataFundacao').send_keys(data_fundacao)
    sleep(2)

    driver.find_element(By.ID,'Cadastrar').click()
    sleep(3)

