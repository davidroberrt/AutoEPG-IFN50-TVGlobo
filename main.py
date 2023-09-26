# Desenvolvido por David Robert de Lima Gomes em 10/03/2023
"""
Este software avançado foi criado para gerar arquivos XML compatíveis com a máquina IFN 50 da Showcase
e automatizar o processo de inserção do Guia Eletrônico de Programação (EPG) 
na Rede Globo da Paraíba (TV Cabo Branco e TV Paraíba). Utilizando bibliotecas como Selenium, o programa interage com a interface web da IFN 50, 
realizando desde o login até a importação e publicação do EPG. Isso simplifica e agiliza consideravelmente a tarefa complexa 
de fornecer informações precisas de programação para transmissão televisiva das duas emissoras filiais da Rede Globo na Paraíba.
"""

import csv
import os
import time
import psutil
import pyautogui
import functools
import xml.etree.ElementTree as ET
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options as FirefoxOptions

def main():
	try:
		# Restante do código aparecendo a 

		# Obtendo data e hora atual
		date_and_time_today = datetime.now()
		date_and_time_today_TEXT = date_and_time_today.strftime('%d%m%Y%H%M%S')
		date_today = datetime.now()
		date = date_today.strftime('%Y%m%d')

		# Obtendo caminhos dos diretórios e arquivos
		path = R"E:\FISCALIZACAO\EPG_TVCBO\EPG_TVCB.{}.xml".format(date) # LOCAL ONDE VAI SALVAR O ARQUIVO FEITO EM XML 
		path_directoryfile = os.path.expandvars(path)
		path2 = R"E:\FISCALIZACAO\EPG_TVCBO" #coloque aqui o diretorio da pasta onde encontra o EPG
		path_directoryfolder = os.path.expandvars(path2)

		path_arquivocsv = R'{}\IN{}_HBCE0.csv'.format(path_directoryfolder,date)

		file_today = R'IN{}_HBCE0.csv'.format(date)

		lista_arquivos = []

		# Listando arquivos na pasta
		pasta = path_directoryfolder

		for diretorio, subpastas, lista_arquivosuivos in os.walk(pasta):
			for lista_arquivosuivo in lista_arquivosuivos:
				a = (os.path.join(lista_arquivosuivo))
				lista_arquivos.append([a])
			
		# Verificando se o arquivo de hoje existe na lista de arquivos
		for i in range(len(lista_arquivos)):
			if '{}'.format(file_today) in lista_arquivos[i]:
				filename = open(path_arquivocsv, 'r')
			
				# Lendo o arquivo CSV
				file = csv.DictReader(filename)
				title = []
				low_description = []
				long_description = []
				broadcast_start_date = []
				broadcast_start_time = []
				seconds = []
				rating = []
				eventID = []

				# Iterando sobre cada linha do arquivo e preenchendo as listas
				for col in file:
					title.append(col['Program title'])
					low_description.append(col['Program content'])
					long_description.append(col['Program content'])
					broadcast_start_date.append(col['Broadcast starting date'])
					broadcast_start_time.append(col['Broadcast starting time'])
					seconds.append(col['Duration'])
					rating.append(col['rating'])
					eventID.append(col['Event ID'])

					quantidadeID = len(eventID) # Quantidade de linhas
				
				# Criando elemento raiz do XML
				data = ET.Element('tv')
				data.set('source_info_url','http://www.showcasepro.tv')
				data.set('source_info_name','ShowCase PRO XMLTV')
				data.set('generator_info_url','http://www.showcasepro.tv/xmltv')
				data.set('generator_info_name','swc-xmltv')
				data.set('date', date_and_time_today_TEXT)

				# Adicionando informações de canais
				channel = ET.SubElement(data, 'channel') 
				channel.set('id', '40640')
				s_nameChannel = ET.SubElement(channel, 'display-name') 
				s_nameChannel.text = "TV CABO BRANCO HD"


				channel2 = ET.SubElement(data, 'channel') 
				channel2.set('id', '40664')
				s_nameChannel = ET.SubElement(channel2, 'display_name')
				s_nameChannel.text = "TV CABO BRANCO 1SEG"
				
				count = 0
				while count < quantidadeID:
				# Adicionando informações dos programas HD
					seconds_int = int(seconds[count])
					programme = ET.SubElement(data, 'programme')
					programme.set("channel", '40640')
					programme.set("start", (broadcast_start_date[count]+broadcast_start_time[count]))

					s_title = ET.SubElement(programme, "title")
					s_title.text = title[count]
					s_desc = ET.SubElement(programme, "desc")
					s_desc.text = low_description[count]
					s_desc_long = ET.SubElement(programme, "desc")
					s_desc_long.text = '...'
					s_category = ET.SubElement(programme, "category")
					s_category.text = "0x8f"
					s_language = ET.SubElement(programme, "language")
					s_language.text = "Portuguese"
					s_length = ET.SubElement(programme, "length")
					s_length.set("units", "seconds")
					s_length.text = seconds[count]
					s_audio = ET.SubElement(programme, "audio")
					s_stereo = ET.SubElement(s_audio, "stereo")
					s_stereo.text = "stereo"
					s_subtitles = ET.SubElement(programme, "subtitles")
					s_subtitles.set("type", "teletext")
					s_language = ET.SubElement(s_subtitles, "language")
					s_language.text = "Portuguese"
					s_rating = ET.SubElement(programme, "rating")
					s_rating.set("system", "SBTVD")
					s_value = ET.SubElement(s_rating, "value")
					s_value.text = rating[count]

					# Adicionando informações dos programas em padrão 1SEG "ONE SEG"

					programme2 = ET.SubElement(data, 'programme')
					programme2.set("channel", '40664')
					programme2.set("start", (broadcast_start_date[count]+broadcast_start_time[count]))

					s_title2 = ET.SubElement(programme2, "title")
					s_title2.text = title[count]
					s_desc2 = ET.SubElement(programme2, "desc")
					s_desc2.text = low_description[count]
					s_desc_long2 = ET.SubElement(programme2, "desc")
					s_desc_long2.text = '...'
					s_category2 = ET.SubElement(programme2, "category")
					s_category2.text = "0x8f"
					s_language2 = ET.SubElement(programme2, "language")
					s_language2.text = "Portuguese"
					s_length2 = ET.SubElement(programme2, "length")
					s_length2.set("units", "seconds")
					s_length2.text = seconds[count]
					s_audio2 = ET.SubElement(programme2, "audio")
					s_stereo2 = ET.SubElement(s_audio2, "stereo")
					s_stereo2.text = "stereo"
					s_subtitles2 = ET.SubElement(programme2, "subtitles")
					s_subtitles2.set("type", "teletext")
					s_language2 = ET.SubElement(s_subtitles2, "language")
					s_language2.text = "Portuguese"
					s_rating2 = ET.SubElement(programme2, "rating")
					s_rating2.set("system", "SBTVD")
					s_value2 = ET.SubElement(s_rating2, "value")
					s_value2.text = rating[count]

					count += 1
				
				# Convertendo o XML em bytes e salvando em um arquivo
				b_xml = ET.tostring(data, encoding='iso-8859-15') 
				with open("{}".format(path_directoryfile), "wb") as f: 
					f.write(b_xml)
					f.close()


				# Configuração de execução do driver Firefox
				flag = 0x08000000  # Sinalizador sem janela
				webdriver.common.service.subprocess.Popen = functools.partial(
				webdriver.common.service.subprocess.Popen, creationflags=flag)
				options = webdriver.FirefoxOptions()
				# Definindo o parâmetro headless (sem interface gráfica)		
				options.headless = True
				driver = webdriver.Firefox(executable_path="C:\geckodriver.exe", options=options)
				driver.implicitly_wait(0.8)
				#driver.set_window_position(-3000, 0) & driver.set_window_position(0, 0) para trazê-lo de volta
				
				# Abrindo o IFN 50
				driver.get("http://IP-da-maquina") #Exclui o IP por segurança, aqui é o link das configurações do encoder.
				assert "IFN" in driver.title

				# Preenchendo informações de login
				textbox_login = driver.find_element(By.ID, "username")
				textbox_login.clear()
				textbox_login.send_keys("admin")
				textbox_password = driver.find_element(By.ID, "password")
				textbox_password.clear()
				time.sleep(0.5)
				textbox_password.send_keys("password")
				time.sleep(0.5)
				# Enviando o formulário de login(logando)
				button_submit = driver.find_element(By.XPATH, '//*[@id="login_form"]/div/div[4]/input').click()
				time.sleep(0.5)

				# Iniciando a importação manual do EPG
				button_import = driver.find_element(By.ID, 'bt_import').click()
				time.sleep(0.5)
				button_manual = driver.find_element(By.XPATH,'/html/body/div[45]/div[2]/div/table/tbody/tr[1]/td/div/label[2]').click()
				time.sleep(0.5)
				button_browse = driver.find_element(By.XPATH,'/html/body/div[45]/div[2]/div/table/tbody/tr[3]/td/div/div[1]/div[1]/div/span[1]').click()
				time.sleep(0.5)

				# Preenchendo o caminho do arquivo XML e realizando a importação
				pyautogui.PAUSE = 1
				pyautogui.typewrite('{}'.format(path_directoryfile))
				time.sleep(1)
				pyautogui.press(['enter'])
				time.sleep(0.5)
				button_epg_import = driver.find_element(By.ID, 'bt_do_epg_import').click()
				time.sleep(0.5)
				button_ok = driver.find_element(By.XPATH,'/html/body/div[46]/div[3]/div/button').click()
				time.sleep(0.5)

				# Publicando o EPG importado
				button_publish = driver.find_element(By.ID, 'bt_publish').click()
				time.sleep(0.5)
				button_publishconfirm = driver.find_element(By.XPATH,'/html/body/div[43]/div[11]/div/button[1]').click()
				filename.close() # Fechando o arquivo CSV
				time.sleep(10)

				# Removendo o arquivo XML gerado após a importação
				if os.path.exists("{}".format(path_directoryfile)):
					os.remove("{}".format(path_directoryfile))

				driver.close()

				# Encerrando processos do computador relacionados(caso precise)
				PROCNAME1 = "geckodriver.exe"
				PROCNAME2 = "Firefox" 
				PROCNAME3 = "{}".format(path_directoryfile) 
				PROCNAME4 = "main" 

				for proc in psutil.process_iter():
					# Verificando se o nome do processo corresponde
					if proc.name() == PROCNAME1:
						proc.kill()
					if proc.name() == PROCNAME2:
						proc.kill()
					if proc.name() == PROCNAME3:
						proc.kill()
					if proc.name() == PROCNAME4:
						proc.kill()
	except Exception as e:
		print("Erro:", e)
        # Exibir um GIF de erro ou outra mensagem de erro
        # Você pode adicionar código para exibir o GIF de erro aqui
if __name__ == "__main__":
    main()