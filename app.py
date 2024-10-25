from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from selenium.webdriver.support.select import Select
import openpyxl

numeroOab = 259155
planilhaDadosConsulta = openpyxl.load_workbook('dados_de_processos.xlsx')
paginaProcessos = planilhaDadosConsulta['processos']

driver = webdriver.Chrome()
driver.get('https://pje-consulta-publica.tjmg.jus.br/')
sleep(5)

campoNumeroOab = driver.find_element(By.XPATH,"//input[@id='fPP:Decoration:numeroOAB']")
sleep(2)
campoNumeroOab.click()
sleep(1)
campoNumeroOab.send_keys(numeroOab)
sleep(1)

selecaoUf = driver.find_element(By.XPATH,"//select[@id='fPP:Decoration:estadoComboOAB']")
sleep(1)
opcoesUf = Select(selecaoUf)
sleep(1)
opcoesUf.select_by_visible_text('SP')
sleep(1)

botaoPesquisar = driver.find_element(By.XPATH,"//input[@id='fPP:searchProcessos']")
sleep(1)
botaoPesquisar.click()
sleep(5)

linksAbrirProcessos = driver.find_elements(By.XPATH,"//a[@title='Ver Detalhes']")

for link in linksAbrirProcessos:
  janelaPrinciapl = driver.current_window_handle
  link.click()
  sleep(5)
  janelasAberta = driver.window_handles
  for janela in janelasAberta:
    if janela not in janelaPrinciapl:
      driver.switch_to.window(janela)
      sleep(5)
      numeroProcesso = driver.find_element(By.XPATH,"//div[@class='propertyView ']//div[@class='col-sm-12 ']")
      participantes = driver.find_elements(By.XPATH,"//tbody[contains(@id,':processoPartesPoloAtivoResumidoList:tb')]//span[@class='text-bold']")
      listaParticipantes = []
      for participante in participantes:
        listaParticipantes.append(participante.text)

      if len(listaParticipantes) == 1:
        paginaProcessos.append([numeroOab, numeroProcesso.text, listaParticipantes[0]])
      else:
        paginaProcessos.append([numeroOab, numeroProcess.text, ','.join(listaParticipantes)])
      planilhaDadosConsulta.save('dados_de_processos.xlsx')
      driver.close()
  driver.switch_to.window(janelaPrinciapl)