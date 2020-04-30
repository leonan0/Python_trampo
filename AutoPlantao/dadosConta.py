from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep
from login import usuario, senha

class usuarios():

    def __init__(self):
        self.driver = webdriver.Chrome()

    def login():
            self.driver.get('http://nt888/portal/Usuario/LogOn')
            print('Abrindo site')
            sleep(8)
            print('esperando')

            apelido = self.driver.find_element_by_xpath('//*[@id="Apelido"]')
            _senha = self.driver.find_element_by_xpath('//*[@id="Senha"]')
            apelido.send_keys(usuario)
            _senha.send_keys(senha)
            print('Logando com Apelido: '+usuario+' e senha: '+senha)

            submit = self.driver.find_element_by_xpath('/html/body/div[1]/form/section/div[2]/input')
            submit.click()
            print('Clicando em ACESSAR')