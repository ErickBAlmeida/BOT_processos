import os
import re
import time

from pywinauto.application import Application
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from win10toast import ToastNotifier

from dict import *


class App:
    def __init__(self):

        self.notifier = ToastNotifier()        

        self.estado = ...    # AQUI VAI A UF DO ESTADO DESEJADO
        
        #   Configurar navegador
        self.navegador = webdriver.Chrome()
        self.navegador.get(uf_link[self.estado])
        self.navegador.maximize_window()

        self.caminho_peticao = os.path.join(os.path.expanduser("~"), "Desktop", "PETICOES")
        self.caminho_procuracao = caminho_procuracao

    def run(self):

        print("\n\nIniciando Automação!\n\n")
        
        self.navegar()
        self.listar_documentos()
    
    def navegar(self):
        
        time.sleep(5)
        self.navegador.refresh()

        #   NAVEGA PELO MENU DO SITE
        try:
            botao_menu = WebDriverWait(self.navegador, 15).until(
                EC.presence_of_element_located((By.CLASS_NAME, "botao-menu"))
            )
            botao_menu.click()
            print("\nNavegando pelo site...")
            time.sleep(1)  # Espera um pouco para garantir que o menu esteja aberto

            processo = self.navegador.find_element("partial link text", "Processo")
            processo.click()
            time.sleep(.5)

            acoes = self.navegador.find_elements(By.XPATH, "//*[contains(text(), ' Outras ações ')]")	
            acoes[0].click()
            time.sleep(.5)

            soliciar_hab = self.navegador.find_elements(By.XPATH, "//*[contains(text(), ' Solicitar habilitação ')]")
            soliciar_hab[0].click()
            time.sleep(.5)
            
        except Exception as e:

            self.notifier.show_toast("ERRO", "Erro ao navegar pelo site", duration=3)
            time.sleep(5)
            raise Exception("Erro ao navegar pelo site") from e
    
    def listar_documentos(self):

        caminho = self.caminho_peticao

        if not os.path.exists(caminho):
            self.notifier.show_toast("ERRO", "O diretório das petições não existe.", duration=3)
            time.sleep(5)
            raise Exception(f"O diretório '{caminho}' não existe.") from e
        
        indice = 0

        for nome_arquivo in os.listdir(caminho):
            caminho_arquivo = os.path.join(caminho, nome_arquivo)
            
            if os.path.isfile(caminho_arquivo):
                
                if self.getNumPeticao(nome_arquivo) == False:
                    continue

                if self.interagirComCheckboxes() == False:
                    continue

                peticao = os.path.join(caminho, nome_arquivo)
                self.incluirDocumentos(peticao, self.caminho_procuracao)
                indice += 1
        
        self.moverPeticao(indice)

        print("\n\nNão há petições, FIM DA AUTOMAÇÃO!!!\n\n\n")
            
    def moverPeticao (self, turnos):
        
        #   Mover petições anexadas para um pasta na área de trabalho ao fim da automação
        for i in range(turnos):
            origem_dir = self.caminho_peticao
            destino_dir = os.path.join(os.path.expanduser("~"), "Desktop", f"{self.estado}_MOVIDAS")
            os.makedirs(destino_dir, exist_ok=True)

            arquivos = [f for f in os.listdir(origem_dir) if os.path.isfile(os.path.join(origem_dir, f))]
            if not arquivos:
                print(f"\nNenhum arquivo encontrado em {origem_dir}")
                return

            primeiro_arquivo = arquivos[0]
            caminho_origem = os.path.join(origem_dir, primeiro_arquivo)
            caminho_destino = os.path.join(destino_dir, primeiro_arquivo)
            os.rename(caminho_origem, caminho_destino)
            print(f"'{primeiro_arquivo}' movido para '{destino_dir}'")

    def getNumPeticao(self, num_peticao):

        try:
            peticao_str = str(num_peticao)

            self.notifier.show_toast("Início", f"Buscando por: \n{peticao_str}", duration=4)
            print("====================================================================")
            print(f"Buscando por: {num_peticao}\n\n")

            peticao_str = re.sub(r'[^0-9]','', peticao_str)

            if peticao_str.isdigit() != True:
                print("\n\nPetição não é valida, indo para a proxima.\n\n")
                return False

            time.sleep(1)
            self.navegador.find_element(By.ID, "fPP:numeroProcesso:numeroSequencial").send_keys(peticao_str[0:7])
            self.navegador.find_element(By.ID, "fPP:numeroProcesso:numeroDigitoVerificador").send_keys(peticao_str[7:9])
            self.navegador.find_element(By.ID, "fPP:numeroProcesso:Ano").send_keys(peticao_str[9:13])
            self.navegador.find_element(By.ID, "fPP:numeroProcesso:NumeroOrgaoJustica").send_keys(peticao_str[-4:])

            # Pesquisar Processo
            time.sleep(1)
            self.navegador.find_element(By.ID, "fPP:searchProcessos").click()

            time.sleep(5)
            try:
                self.navegador.find_element(By.CLASS_NAME, "rich-messages-label")
                print("O processo pesquisado é sigiloso\n\n")
                time.sleep(2)
                self.navegador.find_element("name", "fPP:clearButtonProcessos").click()
                return False

            except:
                print("Abrindo processo...")

            # Solicitar habilitação
            time.sleep(1)
            self.navegador.find_element(By.CLASS_NAME, "botao-link").click()

        except Exception as e:

            self.notifier.show_toast("ERRO", "Erro ao pesquisar por Petição", duration=3)
            time.sleep(5)
            print("Erro ao pesquisar petição\n")
            raise

    def interagirComCheckboxes(self):

        #   MOVER PARA A TERCEIRA ABA
        time.sleep(1)
        abas = self.navegador.window_handles
        self.navegador.switch_to.window(abas[1])

        #   SELECIONAR POLO        
        try:

            time.sleep(1)
            elemento = self.navegador.find_element(By.CLASS_NAME, "btn-default")
            self.navegador.execute_script("arguments[0].scrollIntoView({block: 'center'})", elemento)

            self.navegador.find_element(By.XPATH, "//*[contains(@id, 'selecaoBoxPoloAtivo')]").click() # checkbox 1
            time.sleep(1)
            elemento.click() # checkbox 2

            time.sleep(1)

        except Exception as e:

            self.notifier.show_toast("ERRO", "Não foi possível selecionar polo.", duration=3)
            time.sleep(5)
            print("Erro ao selecionar polo.")
            raise

        #   VINCULAR PARTES        
        try:

            # checkbox = self.navegador.find_element(By.XPATH, "//input[@type='radio' and contains(@id='tipoSolicitacaoDecoration')]")
            checkbox = self.navegador.find_elements(By.XPATH, "//input[@type='radio']")
            self.navegador.execute_script("arguments[0].scrollIntoView({block: 'center'})", checkbox[0])
            
            time.sleep(.5)
            checkbox[0].click()

            time.sleep(1)
            prox_doc = self.navegador.find_element(By.CLASS_NAME,"btn-default")
            self.navegador.execute_script("arguments[0].scrollIntoView({block: 'center'})", prox_doc)

            self.navegador.find_element(By.XPATH, "//input[contains(@name, 'selecaoBoxPoloAtivo')]").click()

        except Exception:

            self.notifier.show_toast("ERRO", "Não foi possível atualizar as partes.", duration=3)
            time.sleep(5)
            raise Exception("Erro ao atualizar as partes.")


        lixeiras = self.navegador.find_elements(By.CLASS_NAME, "btn-sm")
        print(f"Total de partes a serem removidas: {len(lixeiras)}")

        try:
            self.navegador.find_element(By.XPATH, f"//*[contains(text(),{nome_da_parte})]")

            print(f"\n{nome_advogado} já parte desse processo, seguindo para o próximo...\n")
            self.notifier.show_toast("ERRO", f"{nome_advogado} já é parte do processo...", duration=3)

            time.sleep(2)
            self.navegador.close()

            abas = self.navegador.window_handles
            self.navegador.switch_to.window(abas[0])

            time.sleep(1)
            self.navegador.find_element("name", "fPP:clearButtonProcessos").click()

            return False
        
        except:
            pass

        try:    #   REMOVE AS PARTES ANTERIORES

            for i in range(len(lixeiras)):

                print(f"\nExcluindo parte N° {i+1}") 
                time.sleep(.5)
                self.navegador.find_element(By.XPATH, f"//*[contains(@name, '{i}::linkRemovePartePoloAtivo')]").click()

                time.sleep(.5)
                alert = self.navegador.switch_to.alert
                alert.accept()

                print("Parte apagada.\n")
                time.sleep(.5)
            
            time.sleep(2)
            print("Todas as partes removidas.")

        except Exception as e:

            self.notifier.show_toast("ERRO", "Não foi possível excluir partes", duration=3)
            time.sleep(5)
            print(f"ERRO ao excluir partes.\n{e}\n")
        
        self.navegador.find_element(By.XPATH, f"//input[contains(@id, 'tipoDeclaracaoDecoration:tipoDeclaracao:0')]").click()
        self.navegador.find_element(By.XPATH, f"//input[contains(@name, 'botaoProxDoc')]").click()

    def incluirDocumentos(self, caminho_peticao, caminho_procuracao):

        time.sleep(1)
        
        try: #   INCLUIR PETIÇÃO E DOCUMENTOS

            select_option = self.navegador.find_element(By.NAME, "cbTDDecoration:cbTD")
            self.navegador.execute_script("arguments[0].scrollIntoView({block: 'center'})", select_option)
            select = Select(select_option)
            select.select_by_index(index_peticao[self.estado])
            print("Documento do tipo Petição selecionado...")
            time.sleep(1)

        except Exception as e:

            self.notifier.show_toast("ERRO", "Não foi possível selecionar tipo de documento.", duration=3)
            time.sleep(5)
            raise Exception("Erro ao selecionar tipo de documento.") from e
        
        try:
            self.navegador.find_element(By.ID, "raTipoDocPrincipal:0").click()
            time.sleep(3)
            self.navegador.find_element(By.CLASS_NAME, "rich-fileupload-hidden").send_keys(caminho_peticao) # Subir PDF da Petição
            print("Petição anexada.\n")
            
        except:

            self.notifier.show_toast("ERRO", "Não foi possivel subir documento da petição.", duration=3)
            time.sleep(5)
            raise Exception("Erro ao subir petição.") from e

        try:            

            time.sleep(2)
            print("Anexando Procuração...")
            self.navegador.find_element(By.ID, "commandLinkAdicionar").click()      # abre a janela de upload de arquivos

            time.sleep(3)
            
            try:
                app = Application().connect(title="Abrir")      # conecta cm a janela de upload de arquivos

            except Exception as e:
                self.notifier.show_toast("ERRO", "Não foi possivel se conectar ao explorador de arquivos.\nProcesso interrompido!", duration=3)
                time.sleep(5)
                raise Exception("ERRO AO SE CONECTAR COM O EXPLORADOR DE ARQUIVOS.") from e
            
            try:
                dialog = app.window(title_re="Abrir")
                dialog.wait('exists enabled visible ready', timeout=10)
                dialog.set_focus()
                time.sleep(1)
                dialog.type_keys(caminho_procuracao, pause=0.08)
                dialog.type_keys("{ENTER}")
                print("Procuração anexada!\n")

            except Exception as e:
                self.notifier.show_toast("ERRO", "Não foi possível subir o arquivo da procuração.", duration=3)
                time.sleep(5)
                raise Exception("Erro ao tentar subir arquivo da procuração.") from e

        except Exception as e:

            self.notifier.show_toast("ERRO", "Não foi possível realizar o upload da procuração.", duration=3)
            time.sleep(5)
            raise Exception("Input de procuração não localizado.") from e
        
        finally:
            
            print("Assinando petição...")
            
            # ASSINAR PETIÇÃO
            time.sleep(3)
            btn_assinador = self.navegador.find_element(By.ID, "btn-assinador")
            self.navegador.execute_script("arguments[0].scrollIntoView({block: 'center'})", btn_assinador)

            doc_type = WebDriverWait(self.navegador, 5).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(@name, ':tipoDoc')]"))
            )
            
            select = Select(doc_type)
            select.select_by_value(value_procuracao[self.estado])

            time.sleep(8)

            botao_assinador = WebDriverWait(self.navegador, 30).until(
                EC.element_to_be_clickable((By.ID, 'btn-assinador'))
            )

            botao_assinador.click()
            print("assinando...")

            fim = WebDriverWait(self.navegador, 30).until(
                EC.visibility_of_element_located((By.ID, "okBtnConfirmacao"))
            )

            fim.click()
                        
            print("Assinatura concluída!!\n\nFim da rotina...\n")

            time.sleep(1)

            abas = self.navegador.window_handles
            self.navegador.switch_to.window(abas[0])
            time.sleep(1)

            self.navegador.find_element("name", "fPP:clearButtonProcessos").click()

            print("\nReiniciando automação...\n\n===============================================================")

try:
    app = App()
    app.run()
    
except Exception as e:
    print(f"\n\nERRO AO EXECUTAR AUTOMAÇÃO\n\n")
    time.sleep(5)