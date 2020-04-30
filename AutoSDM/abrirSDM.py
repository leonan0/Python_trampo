from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from time import sleep
from datetime import datetime
import json 
import dados
import os.path

class abreSDM():

    def __init__(self):
        self.driver = webdriver.Chrome()

    def logar(self,texto):
        try:
            if os.path.exists('log.txt'):
                log = open('log.txt', 'a')
                log.write(dados.log_padrao+texto)
                print(texto)
            else:
                log = open('log.txt', 'w')
                self.logar(texto)
        except Exception:
            print(Exception)

    def scoreboard(self,relatorio):
        try:
            sleep(90)
            self.troca_frame('product',0)
            self.troca_frame('tab_2001',0)
            self.troca_frame('role_main',0)
            self.troca_frame('scoreboard',0)
            self.driver.implicitly_wait(3)
            x = self.driver.find_element_by_xpath('//span[contains(text(), "{0}")]'.format(relatorio))
            x.click()
            self.logar('Listando ' + relatorio)
            self.frames()
        except:
            self.frames()
            self.logar('Score board - tentando novamente '+relatorio)
            self.scoreboard(relatorio)
            #x = self.driver.find_element_by_xpath('//*[@id="s45pm"]')
            #x = self.driver.find_element_by_xpath('//*[@id="s58ds"]')
            #x = self.drive.find_element_by_xpath('//*[@id="1"]')
            x = self.driver.find_element_by_xpath('//*[@tabindex="3"]')

    def all_incs(self):
        try:
            self.troca_frame('product',0)
            self.troca_frame('tab_2001',0)
            self.troca_frame('role_main',0)
            self.troca_frame('scoreboard',0)
            x = self.driver.find_element_by_xpath('//*[@id="s2pm"]')
            x.click()
            y = self.driver.find_element_by_xpath('//*[@id="s3ds"]')
            y.click()
        except:
            print('123')

            for xs in x:
                y = xs.find_element_by_xpath('//')

 #x = self.driver.find_elements_by_xpath('//*[@tabindex="3"]')


    def exportar(self,export):
        try:
            self.troca_frame('product',0)
            self.troca_frame('tab_2001',0)
            self.troca_frame('role_main',0)
            self.troca_frame('cai_main',0)
            self.driver.implicitly_wait(3)
            x = self.driver.find_element_by_xpath('//span[contains(text(), "{0}")]'.format(export))
            x.click()
            self.logar('Exportando ' +export)
            self.frames()
        except:
            self.frames()
            self.logar('Exportar - tentando novamente '+export)
            self.exportar(export)

    def troca_frame(self,frame_novo,tempo):
        try:
            self.driver.implicitly_wait(tempo)
            self.driver.switch_to.frame(frame_novo)
            self.logar('Novo frame ' + frame_novo)            
        except:
            self.logar('Tentando novamente frame: '+ frame_novo)
            self.troca_frame(frame_novo,3)
            

    def frames(self):
        self.driver.switch_to.default_content()
        self.logar('Frame principal')

    def login(self):
        try:
            self.driver.get('http://'+dados.usuario+':'+dados.senha+'@portosdm/CAisd/pdmweb.exe')
            self.logar('acessando')
        except:
            self.logar('Demorando')
        

        
    #a = self.driver.find_element_by_xpath('//*[@id="imgBtn3"]')

bot = abreSDM()
bot.login()
bot.all_incs()
bot.scoreboard('Incidentes')
bot.scoreboard('Incidentes do meu grupo')
#bot.scoreboard('Minhas pesquisas salvas')
#bot.scoreboard('ALL TCS')
#x = self.driver.find_element_by_xpath('//*[@id="s2pm"]')
#x.click()
#y = self.driver.find_element_by_xpath('//*[@id="s3ds"]')
#y.click()

'''
a = self.driver.find_element_by_id("dataGrid")
itens = a.find_elements_by_tag_name("tr")
for item in itens:
    incidentes = item.find_elements_by_tag_name('td')
    for incidente in incidentes:
        self.logar(incidente.title.text)

p = self.driver.find_element_by_xpath('//*[@id="1"]')

for item in itens:
    incidentes = item.find_elements_by_tag_name('td')
    for incidente in incidentes:
        if incidente.text != '':
            print(incidente.text)'''
