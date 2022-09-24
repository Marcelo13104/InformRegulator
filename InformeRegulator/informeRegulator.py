from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from playsound import playsound
import time
import pandas as pd
import ctypes
import pyperclip as pc





tabela = pd.read_excel("Planilha de informe.xlsx") #Importa a planilha para o codigo ##########################################################

Linha = 0
def Mbox(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)
try:
    navegador = webdriver.Chrome()
    time.sleep(2)
    navegador.get("http://www2.aneel.gov.br/scg/gd/login.asp")
    time.sleep(1)
################################################### Pagina de Login ##########################################################
    #login
    time.sleep(1)
    element = navegador.find_element(By.XPATH,'//*[@id="form"]/table/tbody/tr[2]/td[2]/p/input')
    element.send_keys("regulatorio@elektro.com.br")
    time.sleep(1)
    #senha
    element = navegador.find_element(By.XPATH,'//*[@id="form"]/table/tbody/tr[3]/td[2]/p/input')
    element.send_keys("S88IS972")
    time.sleep(1)
    #logar
    navegador.find_element(By.XPATH,'//*[@id="form"]/table/tbody/tr[4]/td/p/input[2]').click()
    time.sleep(1)

    
################################################### Primeira Pagina ##########################################################
    def Cadastro(Linha):
            #inicia o cadastro de um nova UC GD
            navegador.find_element(By.XPATH,'//*[@id="form"]/table/tbody/tr[2]/td[2]/input').click()
            time.sleep(1)
            #Unidade Consumidora
            element = navegador.find_element(By.XPATH,'//*[@id="form1"]/table/tbody/tr[2]/td/input')
            element.send_keys(str(tabela.loc[Linha,"UC"]))
            time.sleep(1)
            #Campo de selecionar Modalidade
            select = Select(navegador.find_element(By.NAME, 'IdcModalidade'))
            select.select_by_visible_text(str(tabela.loc[Linha,"Modalidade"]))
            time.sleep(1)
            # Quantidade de Unidades Consumidoras que recebem creditos
            navegador.find_element(By.XPATH,'//*[@id="form1"]/table/tbody/tr[3]/td[2]/input').clear()
            element = navegador.find_element(By.XPATH,'//*[@id="form1"]/table/tbody/tr[3]/td[2]/input')
            element.send_keys(str(tabela.loc[Linha,"Quantidade Beneficiarias"]))
            time.sleep(1)
            #Campo de selecionar classe de fornecimento
            select = Select(navegador.find_element(By.NAME, 'IdeClasseFornecimento'))
            select.select_by_visible_text(str(tabela.loc[Linha,"Classe"]))
            time.sleep(1)
            #Campo de selecionar SubGrupo
            select = Select(navegador.find_element(By.NAME, 'IdeGrupoFornecimento'))
            select.select_by_visible_text(str(tabela.loc[Linha,"Grupo Tensão"]))
            time.sleep(1)
            #Campo de selecionar cidade
            try:
                element = navegador.find_element(By.XPATH,'//*[@id="NomeBusca1"]')
                element.send_keys(str(tabela.loc[Linha,"Municipio"]))
                time.sleep(3)
                codMunicipio = int(tabela.loc[Linha,"CodMunicipio"])
                #selectCidade = (str(tabela.loc[Linha,"Municipio"]) +" / SP")
                select = Select(navegador.find_element(By.ID, 'IdSelecionado1'))
                select.select_by_value(str(codMunicipio))
                time.sleep(1)

            except:
                Mbox('Erro', 'Erro ao selecionar municipio de '+ str(tabela.loc[Linha,"Municipio"]), 0)
                 


            #Endereço
            element = navegador.find_element(By.XPATH,'//*[@id="form1"]/table/tbody/tr[5]/td[1]/input')
            element.send_keys(str(tabela.loc[Linha,"Endereço"]))
            time.sleep(1)
            #CEP
            cep = str(tabela.loc[Linha,"Cep"])
            if len(cep) == 7:
                cep = "0" + cep
            elif len(cep) == 6:
                cep = "00" + cep
            element = navegador.find_element(By.XPATH,'//*[@id="form1"]/table/tbody/tr[5]/td[2]/input')
            element.send_keys(cep)
            time.sleep(1)
            #Campo de selecionar fonte
            select = Select(navegador.find_element(By.NAME, 'IdeCombustivel'))
            select.select_by_visible_text(str(tabela.loc[Linha,"Fonte"]))
            time.sleep(1)
            #Campo de selecionar coordenadas geograficas
            #latitude
            element = navegador.find_element(By.XPATH,'//*[@id="form1"]/table/tbody/tr[8]/td[1]/select[1]')
            element.send_keys(str(tabela.loc[Linha,"lat grau"]))
            time.sleep(1)

            element = navegador.find_element(By.XPATH,'//*[@id="form1"]/table/tbody/tr[8]/td[1]/select[2]')
            element.send_keys(str(tabela.loc[Linha,"lat min"]))
            time.sleep(1)

            navegador.find_element(By.XPATH,'//*[@id="form1"]/table/tbody/tr[8]/td[1]/input').clear()
            element = navegador.find_element(By.XPATH,'//*[@id="form1"]/table/tbody/tr[8]/td[1]/input')
            time.sleep(1)
            element.send_keys(str(tabela.loc[Linha,"lat seg"]))
            time.sleep(1)
            #longitude
            element = navegador.find_element(By.XPATH,'//*[@id="form1"]/table/tbody/tr[8]/td[2]/select[1]')
            element.send_keys(str(tabela.loc[Linha,"lon grau"]))
            time.sleep(1)

            element = navegador.find_element(By.XPATH,'//*[@id="form1"]/table/tbody/tr[8]/td[2]/select[2]')
            element.send_keys(str(tabela.loc[Linha,"lon min"]))
            time.sleep(1)

            navegador.find_element(By.XPATH,'//*[@id="form1"]/table/tbody/tr[8]/td[2]/input').clear()
            element = navegador.find_element(By.XPATH,'//*[@id="form1"]/table/tbody/tr[8]/td[2]/input')
            element.send_keys(str(tabela.loc[Linha,"lon seg"]))
            time.sleep(1)

            # campo inserir o CPF
            cpf = str(tabela.loc[Linha,"CNPJ"])
            if len(cpf) == 10:
                cpf = "0" + cpf
            elif len(cpf) == 9:
                cpf = "00" + cpf
            elif len(cpf) == 8:
                cpf = "000" + cpf
            elif len(cpf) == 7:
                cpf = "0000" + cpf
            elif len(cpf) == 13:
                cpf = "0" + cpf
            elif len(cpf) == 12:
                cpf = "00" + cpf
            element = navegador.find_element(By.XPATH,'//*[@id="form1"]/table/tbody/tr[11]/td[1]/input[1]')
            element.send_keys(cpf)
            time.sleep(1)
            # campo inserir o nome do titular
            element = navegador.find_element(By.XPATH,'//*[@id="form1"]/table/tbody/tr[11]/td[2]/input')
            element.send_keys(str(tabela.loc[Linha,"Nome"]))
            time.sleep(1)
            # campo inserir o telefone
            telefone = str(tabela.loc[Linha,"Telefone "])
            if len(telefone) == 0 or len(telefone) == 1:
                telefone = "11111111111"
            element = navegador.find_element(By.XPATH,'//*[@id="form1"]/table/tbody/tr[12]/td[1]/input')
            element.send_keys(telefone)
            time.sleep(1)
            # campo inserir o E-mail
            email = str(tabela.loc[Linha,"E-mail"])
            if len(email) == 0 or len(email) == 1:
                email = "nao@existe.com"
            element = navegador.find_element(By.XPATH,'//*[@id="form1"]/table/tbody/tr[12]/td[2]/input')
            element.send_keys(email)
            time.sleep(1)

            #Campo de selecionar municipio
            try:
                formatCidade = str(tabela.loc[Linha,"Municipio"])
                if formatCidade == "Mogi Mirim" or "Mogi-Mirim":
                    formatCidade == "Moji-mirim"
                element = navegador.find_element(By.XPATH,'//*[@id="NomeBusca2"]')
                element.send_keys(formatCidade)
                time.sleep(3)
                codMunicipio = int(tabela.loc[Linha,"CodMunicipio"])
                select = Select(navegador.find_element(By.ID, 'IdSelecionado2'))
                select.select_by_value(str(codMunicipio))
                time.sleep(1)
            except:
                Mbox('Erro', 'Erro ao selecionar municipio de '+ str(tabela.loc[Linha,"Municipio"]), 0)
                 
            #Endereço
            element = navegador.find_element(By.XPATH,'//*[@id="form1"]/table/tbody/tr[15]/td[1]/input')
            element.send_keys(str(tabela.loc[Linha,"Endereço"]))
            time.sleep(1)
            #CEP
            cep = str(tabela.loc[Linha,"Cep"])
            if len(cep) == 7:
                cep = "0" + cep
            elif len(cep) == 6:
                cep = "00" + cep
            element = navegador.find_element(By.XPATH,'//*[@id="form1"]/table/tbody/tr[15]/td[2]/input')
            element.send_keys(cep)
            time.sleep(1)
################################################### Segunda Pagina ##########################################################

            navegador.find_element(By.XPATH,'//*[@id="form1"]/table/tbody/tr[16]/td/input[3]').click()
            time.sleep(5)
                
            #Potencia Total dos Modulos
            try:
                
                navegador.find_element(By.XPATH,'/html/body/table[3]/tbody/tr[1]/td[2]/input').clear()
                element = navegador.find_element(By.XPATH,'/html/body/table[3]/tbody/tr[1]/td[2]/input')
                pc.copy(tabela.loc[Linha,"Potencia Modulos"])
                element.send_keys(Keys.CONTROL + "v")
                time.sleep(1)

            except:
                Mbox('Erro', 'Erro no campo Potencia Modulos, Linha '+str(Linha+2),0)
                

            #Quantidade Total dos Modulos
            navegador.find_element(By.XPATH,'/html/body/table[3]/tbody/tr[1]/td[4]/input').clear()
            element = navegador.find_element(By.XPATH,'/html/body/table[3]/tbody/tr[1]/td[4]/input')
            pc.copy(str(tabela.loc[Linha,"Quantidade Modulos"]))
            element.send_keys(Keys.CONTROL + "v")
            time.sleep(1)
            #Potencia Total dos Inversores
            navegador.find_element(By.XPATH,'/html/body/table[3]/tbody/tr[2]/td[2]/input').clear()
            element = navegador.find_element(By.XPATH,'/html/body/table[3]/tbody/tr[2]/td[2]/input')
            pc.copy(tabela.loc[Linha,"Potencia Inversores"])
            element.send_keys(Keys.CONTROL + "v")
            time.sleep(1)

            #Quantidade Total dos Inversores
            try:
                navegador.find_element(By.XPATH,'/html/body/table[3]/tbody/tr[2]/td[4]/input').clear()
                element = navegador.find_element(By.XPATH,'/html/body/table[3]/tbody/tr[2]/td[4]/input')
                pc.copy(str(tabela.loc[Linha,"Quantidade Inversores"]))
                element.send_keys(Keys.CONTROL + "v")
                time.sleep(1)
            except:
                Mbox('Erro', 'Erro no campo Quantidade Total dos Inversores, Linha '+str(Linha+2), 0)

            #Área Total dos Arranjos  
            navegador.find_element(By.XPATH,'/html/body/table[3]/tbody/tr[4]/td[2]/input').clear()
            element = navegador.find_element(By.XPATH,'/html/body/table[3]/tbody/tr[4]/td[2]/input')
            pc.copy(tabela.loc[Linha,"Area Total"])
            element.send_keys(Keys.CONTROL + "v")
            time.sleep(1)

            #Fabricante(s) dos Módulos
            element = navegador.find_element(By.XPATH,'/html/body/table[3]/tbody/tr[5]/td[2]/input')
            element.send_keys(tabela.loc[Linha,"Fabricante Modulos"])
            time.sleep(1)
            #Modelo(s) dos Módulos
            element = navegador.find_element(By.XPATH,'/html/body/table[3]/tbody/tr[5]/td[4]/input')
            element.send_keys(tabela.loc[Linha,"Modelo Modulos"])
            time.sleep(1)
            #Fabricante(s) dos Inversores
            element = navegador.find_element(By.XPATH,'/html/body/table[3]/tbody/tr[6]/td[2]/input')
            element.send_keys(str(tabela.loc[Linha,"Fabricante Inversores"]))
            time.sleep(1)
            #Modelo(s) dos Inversores
            element = navegador.find_element(By.XPATH,'/html/body/table[3]/tbody/tr[6]/td[4]/input')
            element.send_keys(str(tabela.loc[Linha,"Modelo Inversores"]))
            time.sleep(1)
            #Data da implantação
            dataImplatacao = pd.to_datetime(pd.Series(tabela.loc[Linha,"Data Implatação"])) #Formata a data de Implantação
            element = navegador.find_element(By.XPATH,'/html/body/table[3]/tbody/tr[7]/td[2]/input')
            element.send_keys(dataImplatacao.dt.strftime('%d/%m/%Y'))
            time.sleep(1)
            #Data da conexão da GD
            dataConexao = pd.to_datetime(pd.Series(tabela.loc[Linha,"Data Conexão"])) #Formata a data de conexão
            element = navegador.find_element(By.XPATH,'/html/body/table[3]/tbody/tr[7]/td[4]/input')
            element.send_keys(dataConexao.dt.strftime('%d/%m/%Y'))
            time.sleep(1)
            #Proxima Pagina
            navegador.find_element(By.XPATH,'/html/body/table[3]/tbody/tr[8]/td/input[3]').click()
            time.sleep(1)
            #ultima pagina
            navegador.find_element(By.XPATH,'//*[@id="form1"]/table/tbody/tr[4]/td/input[3]').click()
            time.sleep(1)
            # Guarda o numero GD e Salva planilha
            element = navegador.find_element(By.XPATH,'/html/body/table[2]/tbody/tr[1]/td/form/table/tbody/tr[2]/td/p[3]')
            print(element.text)
            #Salva a Planilha
            tabela.loc[Linha,"Numero GD"] = element.text
            tabela.to_excel("Cadastrados.xlsx",index=False)
            #finaliza e volta para a pagina inicio
            navegador.find_element(By.XPATH,'/html/body/table[2]/tbody/tr[1]/td/form/table/tbody/tr[3]/td/input[2]').click()
            time.sleep(1)
except:
    Mbox('Erro', 'Algo deu errado, Verifique os Campos! Linha '+str(Linha+2), 0)
       

################################################### Cadastro em massa ##########################################################
try:
    while Linha <= len(tabela.index):
        if str(tabela.loc[Linha,"Status2"]) == "não cadastrado":
            Cadastro(Linha)
        Linha = Linha + 1
        
    if Linha == len(tabela.index):
        print("Fim!!")
        Mbox('ALERTA', 'Cadastros Finalizados', 0) 
            
except:
    Mbox('Erro', 'Algo deu errado, verifique os Dados da Linha '+str(Linha+2), 0)