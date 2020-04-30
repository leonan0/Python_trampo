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

    def iniciar(self):
        self.login()
        self.all_incs()
        self.listagem()
        self.looper()
        return ''

    def logar(self,texto):
        try:
            if os.path.exists('incidentes'+dados.datah+'.json'):
                log = open('incidentes'+dados.datah+'.json', 'a')
                log.write(texto)
            else:
                log = open('incidentes'+dados.datah+'.json', 'w')
                self.logar(texto)
        except Exception:
            print(Exception)

    def montaestrutura(self, texto, param):
        if param == 1:
	        self.logar('\n		{\n			')
        elif param == 2:
            texto = texto.replace(' ', '')
            texto = texto.replace('*', '')
            self.logar('"numero":"' + texto + '",\n			')
        elif param == 3:
            self.logar('"resumo":"' + texto + '",\n			')
        elif param == 4:
            self.logar('"severidade":"' + texto + '",\n			')
        elif param == 5:
            self.logar('"categoria":"' + texto + '",\n			')
        elif param == 6:
            self.logar('"status":"' + texto + '",\n			')
        elif param == 7:
            self.logar('"grupo_executor":"' + texto + '",\n			')
        elif param == 8:
            self.logar('"responsavel":"' + texto + '",\n			')
        elif param == 9:
            self.logar('"violacao_projetada":"' + texto + '",\n			')
        elif param == 10:
            self.logar('"violado":"' + texto + '",\n			')
        elif param == 11:
            self.logar('"localidade":"' + texto + '",\n			')
        elif param == 12:
            self.logar('"data_abertura":"' + texto + '",\n			')
        elif param == 13:
            self.logar('"ultima_atualizacao":"' + texto + '",\n			')
        elif param == 14:
            self.logar('"retorno_chamado":"' + texto + '",\n			')
        elif param == 15:
            self.logar('"classificacao_final":"' + texto + '",\n			')
        elif param == 16:
            self.logar('"data_resolucao":"' + texto + '",\n			')
        elif param == 17:
            self.logar('"descricao":"' + texto + '",\n			')
        elif param == 18:
            self.logar('"usuario_final_afetado":"' + texto + '",\n			')
        elif param == 19:
            self.logar('"departamento":"' + texto + '",\n			')
        elif param == 20:
            self.logar('"problema_vinculado":"' + texto + '",\n			')
        elif param == 21:
            self.logar('"incidente_pai":"' + texto + '",\n			')
        elif param == 22:
            self.logar('"causado_pela_rdm":"' + texto + '",\n			')
        elif param == 23:
            self.logar('"origem":"' + texto + '",\n			')
        elif param == 24:
            self.logar('"ticket_sis_ext":"' + texto + '"\n		}')
            dados.count = dados.count + 1
            if dados.count != dados.total_inc:
                self.logar(',') 



    def all_incs(self):
        try:
            self.frames() 
            self.troca_frame('product',0)
            self.troca_frame('tab_2001',0)
            self.troca_frame('role_main',0)
            self.troca_frame('scoreboard',0)
            x = self.driver.find_element_by_xpath('//*[@id="s2pm"]')
            x.click()
            y = self.driver.find_element_by_xpath('//*[@id="s3ds"]')
            y.click()
            dados.loop = self.loops() 
        except:
            print('all incs tentando novamente')
            self.all_incs()

    def listagem(self):
        try:
            self.frames()
            self.troca_frame('product',1)
            self.troca_frame('tab_2001',1)
            self.troca_frame('role_main',1)
            self.troca_frame('cai_main',1)
        except:
            print('listagem tentando novamente')
            self.listagem()

    def troca_frame(self,frame_novo,tempo):
        try:
            self.driver.implicitly_wait(tempo)
            self.driver.switch_to.frame(frame_novo)
        except:
            print('troca_fame tentando novamente')
            self.troca_frame(frame_novo,3)
            
    def frames(self):
        self.driver.switch_to.default_content()
        print('frame principal')

    def login(self):
        try:
            self.driver.get('http://'+dados.usuario+':'+dados.senha+'@portosdm/CAisd/pdmweb.exe')
            print('acessando')
        except:
            print('Demorando')
    
    def loops(self):
        dados.total_inc = self.driver.find_element_by_id('s3ct')
        dados.total_inc = int(dados.total_inc.text)
        loop = dados.total_inc / 25
        loop = round(loop)
        return loop

    def looper(self):
        l = dados.loop
        while l >= 1:
            self.teste(l)
            l = l-1
            self.driver.find_element_by_xpath('//*[@id="next_t_dataGrid_toppager"]/a').click()

    def teste(self,pag):
        print('listando pag - ' + str(pag))
        a = 0
        if pag == dados.loop:
            self.logar('{\n\n	"incidentes":[')
        b = 0
        data = self.driver.find_element_by_id('dataGrid')
        datab = data.find_elements_by_tag_name('tr')
        for itens in datab:
            a = a+1
            incidentes = itens.find_elements_by_tag_name('td')
            if a > 2:
                b =1
                for incidente in incidentes:
                    info = incidente.get_attribute('title')
                    info = info.replace('"', '-')
                    self.montaestrutura(info,b)
                    b = b+1 
        if pag == 1:
            self.logar('\n	]\n}')

        
'''
bot = abreSDM()
bot.login()
bot.all_incs()
bot.listagem()
bot.looper()'''