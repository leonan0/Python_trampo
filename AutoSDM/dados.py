import datetime

teste_botao = 'Minhas pesquisas salvas'
minhaspesquisassalvas ='//*[@id="s2pm"]'
minhaspesquisassalvas1 ='//*[@id="s1pm"]'
pesquisa1 = '//*[@id="s4ds"]'
pesquisa2 = '//*[@id="s5ds"]'
usuario =  'P0642065'
senha = 'umdois33'
t = datetime.datetime.now()
databuscaa = t - datetime.timedelta(days=5)
databusca = databuscaa.strftime("%d/%m/%Y")
data = t.strftime("%d/%m/%Y %H:%M:%S")
datah = t.strftime("%d%m%Y %H")
loop = 0
total_inc = 0
count = 0
link = 'http://'+usuario+':'+senha+'@portosdm/CAisd/pdmweb.exe'
scoreboard = ['product','tab_2001','role_main','scoreboard']
cai_main = ['product','tab_2001','role_main','cai_main']
click_scoreboard = [minhaspesquisassalvas1,pesquisa2]
#click_scoreboard = [minhaspesquisassalvas,pesquisa2]
click_scoreboard2 = [pesquisa1]
campoparam = 'sf_7_1'
pesquisar_sdm = '//*[@id="imgBtn0"]'
contador = 'sp_1_dataGrid_toppager'
mostrafiltros = '//*[@id="imgBtn1"]'
maisfiltros = '//*[@id="sf_4_2"]'




