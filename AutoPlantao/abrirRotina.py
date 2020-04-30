from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep
from padroes import log_padrao, usuario, senha
from datetime import datetime


class abreRotina():

    def __init__(self):
        self.driver = webdriver.Chrome()

    def logar(self,texto):
        try:
            log = open('log.txt', 'a')
            log.write(log_padrao+texto)
        except Exception:
            print(Exception)

    def login(self,nt):
        self.driver.get('http://'+nt + '/portal/Usuario/LogOn')
        self.logar('ascessando ' + nt +' para logar')
        
        apelido = self.driver.find_element_by_xpath('//*[@id="Apelido"]')
        _senha = self.driver.find_element_by_xpath('//*[@id="Senha"]')
        apelido.send_keys(usuario)
        _senha.send_keys(senha)

        submit = self.driver.find_element_by_xpath('/html/body/div[1]/form/section/div[2]/input')
        submit.click()
        nome = self.driver.find_element_by_xpath('/html/body/header/div[2]/div[1]').text
        
        self.logar('dados de acesso')
        self.logar(nome)
        self.logar('Usuario: ' + usuario)
        self.logar('Senha: ' + senha)
        sleep(60)
        

bot = abreRotina()
bot.login('nt1510')

