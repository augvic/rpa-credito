# ================================================== #

# ~~ Bibliotecas.
from selenium import webdriver
from selenium.webdriver.chrome.options import Options as opt
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.common.by import By as By
import xlwings as xw
import win32com.client
from datetime import datetime, timedelta
import pandas as pd
import os
import keyboard
import threading

# ================================================== #

# ~~ Classe RPA - Crédito.
class RPACrédito:

    """
    RPA - Crédito
    ---
    Faz análise de crédito dos pedidos do site retornando com sua liberação ou recusa.
    """

    # ================================================== #

    # ~~ Ao criar instância.
    def __init__(self):

        """
        Resumo:
        * Cria log e define variável global de encerramento ao instanciar RPA.
        ---
        Parâmetros:
        * ===
        ---
        Retorna:
        * ===
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """

        # ~~ Cria log.
        self.Log = ""

        # ~~ Define encerramento global.
        self.Encerrar = False

        # ~~ Define reinicio do loop.
        self.ReiniciarLoop = False

        # ~~ Define horário de início.
        self.DataHoraInício = datetime.now().replace(microsecond = 0)
        self.DataHoraInício = self.DataHoraInício.strftime("%d-%m-%Y_%H-%M")

    # ================================================== #

    # ~~ Função customizada para printar mensagens.
    def PrintarMensagem(self, Mensagem: str = None, CharType: str = None, Qtd: int = None, Side: str = None) -> None:

        """
        Resumo:
        * Printa mensagem ou caractere especial no terminal.
        ---
        Parâmetros:
        * Mensagem (opcional) -> Texto que será printado no terminal.
        * CharType (opcional) -> Tipo de caractére a ser printado.
        * Qtd (opcional) -> Quantidade do caractére.
        * Side (opcional) -> Lado que será printado: cima/baixo.
        ---
        Retorna:
        * ===
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """

        # ~~ Coleta data e hora atual.
        DataHoraAtual = datetime.now().replace(microsecond = 0)
        DataHoraAtual = DataHoraAtual.strftime("%d/%m/%Y_%H:%M")

        # ~~ Se tiver Mensagem.
        if Mensagem:

            # ~~ Se tiver Mensagem e CharType, verifica o lado a ser printado.
            if CharType:
                if Side == "top":
                    print(f"<{DataHoraAtual}>")
                    print(CharType*Qtd)
                    print(Mensagem)
                    self.Log += f"<{DataHoraAtual}>\n{CharType*Qtd}\n{Mensagem}\n"
                if Side == "bot":
                    print(f"<{DataHoraAtual}>")
                    print(Mensagem)
                    print(CharType*Qtd)
                    self.Log += f"<{DataHoraAtual}>\n{Mensagem}\n{CharType*Qtd}\n"
                if Side == "both":
                    print(f"<{DataHoraAtual}>")
                    print(CharType*Qtd)
                    print(Mensagem)
                    print(CharType*Qtd)
                    self.Log += f"<{DataHoraAtual}>\n{CharType*Qtd}\n{Mensagem}\n{CharType*Qtd}\n"

            # ~~ Se não tiver CharType, printa somente Mensagem.
            else:
                print(f"<{DataHoraAtual}>")
                print(Mensagem)
        
        # ~~ Se não tiver Mensagem, printa somente CharType.
        else:
            print(CharType*Qtd)

    # ================================================== #

    # ~~ Cria instância do navegador, utilizando o webdriver.
    def InstanciarNavegador(self) -> webdriver:

        """
        Resumo:
        * Cria instância do navegador, utilizando o webdriver.
        ---
        Parâmetros:
        * ===
        ---
        Retorna:
        * Driver -> Navegador instanciado.
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """

        # ~~ Definindo configurações.
        Options = opt()
        Options.add_argument("--log-level=3")
        Options.add_experimental_option("excludeSwitches", ["enable-logging"])
        Options.add_experimental_option("detach", True) 

        # ~~ Criando instância.
        Driver = webdriver.Chrome(options=Options) 
        AbasAbertas = Driver.window_handles
        if len(AbasAbertas) > 1:
            Driver.switch_to.window(AbasAbertas[0])
            Driver.close()
        try:
            Driver.switch_to.window(AbasAbertas[0])
        except:
            Driver.switch_to.window(AbasAbertas[1])

        # ~~ Acessando GoDeep e fazendo login.
        Driver.get(f"https://www.revendedorpositivo.com.br/admin/")
        microsoft_login_botao = None
        try:
            microsoft_login_botao = Driver.find_element(By.ID, value="login-ms-azure-ad")
            microsoft_login_botao.click()
            time.sleep(3)
            body = Driver.find_element(By.TAG_NAME, value="body").text
            if any(login_string in body for login_string in ["Because you're accessing sensitive info, you need to verify your password.", "Sign in", "Pick an account", "Entrar"]):
                self.PrintarMensagem("Necessário logar conta Microsoft.", "=", 50, "bot")
                while True:
                    body = Driver.find_element(By.TAG_NAME, value="body").text
                    if "DASHBOARD" in body:
                        break
                    else:
                        time.sleep(3)
            if "Approve sign in request" in body:
                time.sleep(3)
                codigo = Driver.find_element(By.ID, value="idRichContext_DisplaySign").text
                self.PrintarMensagem(f"Necessário authenticator Microsoft para continuar: {codigo}.", "=", 50, "bot")
                while True:
                    body = Driver.find_element(By.TAG_NAME, value="body").text
                    if "DASHBOARD" in body:
                        break
                    else:
                        time.sleep(3)
        except:
            Driver.get(f"https://www.revendedorpositivo.com.br/admin/index/")

        # ~~ Retorna instância do webdriver.
        return Driver

    # ================================================== #

    # ~~ Cria instância da planilha de controle, utilizando o xlwings.
    def InstanciarControle(self) -> dict:

        """
        Resumo:
        * Cria instância do controle de crédito (planilha do Excel).
        ---
        Parâmetros:
        * ===
        ---
        Retorna:
        * Controle -> Dicionário contendo: ["BOOK"] - ["PEDIDOS"] - ["LIMITES"].
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """

        # ~~ Cria dicionário.
        Controle = {}

        # ~~ Crian instância.
        CaminhoScript = os.path.abspath(__file__)
        CaminhoControle = CaminhoScript.split(r"\script_crédito.py")[0] + r"\controle_crédito.xlsx"
        Controle["BOOK"] = xw.Book(CaminhoControle)
        Controle["PEDIDOS"] = Controle["BOOK"].sheets["PEDIDOS"]
        Controle["LIMITES"] = Controle["BOOK"].sheets["LIMITES"]

        # ~~ Retorna instância do controle.
        return Controle

    # ================================================== #

    # ~~ Cria instância do SAP.
    def InstanciarSap(self) -> object:

        """
        Resumo:
        * Cria instância com o SAP, acessando a SAPScriptingEngine.
        ---
        Parâmetros:
        * ===
        ---
        Retorna:
        * Session -> Conexão com SAP.
        ---
        Erros:
        * Não foi encontrado tela SAP disponível para conexão.
        ---
        Erros tratados localmente:
        * Não foi encontrado tela SAP disponível para conexão.
        ---
        Erros levantados:
        * ===
        """

        # ~~ Tenta conexão.
        try:
            Gui = win32com.client.GetObject("SAPGUI")
            App = Gui.GetScriptingEngine
            Con = App.Children(0)
            for Id in range(0, 4):
                Session = Con.Children(Id)
                if Session.ActiveWindow.Text == "SAP Easy Access":
                    return Session
                else:
                    continue

            # ~~ Se não encontrar tela disponível.
            else:
                self.PrintarMensagem("Não foi encontrado tela SAP disponível para conexão.", "=", 30, "bot")
                exit()
        
        # ~~ Se não encontrar tela logada no SAP.
        except:
            self.PrintarMensagem("Não foi encontrado tela SAP disponível para conexão.", "=", 30, "bot")
            exit()

    # ================================================== #

    # ~~ Acessa pedido no site.
    def AcessarPedido(self, Pedido: int) -> None:

        """
        Resumo:
        * Acessa a página do pedido no site.
        ---
        Parâmetros:
        * Pedido -> Número do pedido.
        ---
        Retorna:
        * ===
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """

        # ~~ Time sleep.
        time.sleep(3)

        # ~~ Acessa pedido.
        self.Driver.get(f"https://www.revendedorpositivo.com.br/admin/orders/edit/id/{Pedido}")

    # ================================================== #

    # ~~ Coleta data do pedido.
    def ColetarDataPedido(self) -> datetime:

        """
        Resumo:
        * Coleta a data do pedido no site.
        ---
        Parâmetros:
        * ===
        ---
        Retorna:
        * Data -> Data do pedido.
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """

        # ~~ Coleta data.
        Data = self.Driver.find_element(By.XPATH, value="//label[@for='order_date']/following-sibling::div[@class='col-md-12']").text
        Data = datetime.strptime(Data, "%d/%m/%Y %H:%M:%S")

        # ~~ Retorna a data.
        return Data

    # ================================================== #

    # ~~ Coleta condição de pagamento do pedido.
    def ColetarCondiçãoPagamento(self) -> str:

        """
        Resumo:
        * Coleta a forma de pagamento do pedido no site.
        ---
        Parâmetros:
        * ===
        ---
        Retorna:
        * FormaPagamento -> Forma de pagamento.
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """

        # ~~ Coleta condição de pagamento.
        FormaPagamento = self.Driver.find_element(By.XPATH, value="//label[@for='payment_slip_installments_description']/following-sibling::div[@class='col-md-12']").text

        # ~~ Retorna a condição de pagamento.
        return FormaPagamento

    # ================================================== #

    # ~~ Coleta forma de pagamento do pedido.
    def ColetarFormaPagamentoPedido(self) -> str:

        """
        Resumo:
        * Coleta a forma de pagamento do pedido no site.
        ---
        Parâmetros:
        * ===
        ---
        Retorna:
        * FormaPagamento -> Forma de pagamento.
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """

        # ~~ Coleta forma de pagamento.
        FormaPagamento = self.Driver.find_element(By.XPATH, value="//label[@for='payment_name']/following-sibling::div[@class='col-md-12']").text

        # ~~ Retorna forma de pagamento.
        return FormaPagamento

    # ================================================== #

    # ~~ Coleta CNPJ do cliente.
    def ColetarCnpj(self) -> str:

        """
        Resumo:
        * Coleta o CNPJ do cliente no site.
        ---
        Parâmetros:
        * ===
        ---
        Retorna:
        * Cnpj -> Número do CNPJ sem caracteres especiais.
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """

        # ~~ Coleta CNPJ.
        Cnpj = self.Driver.find_element(By.XPATH, value="//label[@for='client_cnpj']/following-sibling::div[@class='col-md-12']").text
        Cnpj = Cnpj[:8]

        # ~~ Retorna CNPJ.
        return Cnpj

    # ================================================== #

    # ~~ Coleta código ERP do cliente no SAP.
    def ColetarCódigoERP(self, CNPJ: str) -> str:

        """
        Resumo:
        * Coleta o código ERP do cliente no SAP.
        ---
        Parâmetros:
        * CNPJ -> CNPJ do cliente.
        ---
        Retorna:
        * CódigoERP -> Código ERP do cliente.
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """

        # ~~ Acessa XD03.
        self.AbrirTransação("XD03")

        # ~~ Busca pelo CNPJ.
        self.Session.findById("wnd[1]").sendVKey(4)
        self.Session.findById("wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB006").select()
        self.Session.findById("wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]").text = CNPJ
        self.Session.findById("wnd[2]/tbar[0]/btn[0]").press()
        StatusBarMsg = self.Session.findById("wnd[0]/sbar").text
        if "Nenhum valor para esta seleção" in StatusBarMsg:
            self.Session.findById("wnd[1]").close()
            self.Session.findById("wnd[1]").close()
            return "-"
        self.Session.findById("wnd[2]").sendVKey(2)

        # ~~ Coleta código ERP.
        CódigoERP = self.Session.findById("wnd[1]/usr/ctxtRF02D-KUNNR").text

        # ~~ Fecha transação.
        self.Session.findById("wnd[1]").close()

        # ~~ Retorna código ERP.
        return CódigoERP

    # ================================================== #

    # ~~ Coleta valor do pedido.
    def ColetarValorPedido(self) -> float:

        """
        Resumo:
        * Coleta o valor do pedido no site.
        ---
        Parâmetros:
        * ===
        ---
        Retorna:
        * ValorPedido -> Valor do pedido.
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """

        # ~~ Coleta valor do pedido.
        ValorPedido = self.Driver.find_element(By.XPATH, value="//label[@for='payment_value']/following-sibling::div[@class='col-md-12']").text 
        ValorPedido = ValorPedido.replace("R$", "").replace(".", "").replace(",", ".")
        ValorPedido = float(ValorPedido)

        # ~~ Retorna valor do pedido.
        return ValorPedido

    # ================================================== #

    # ~~ Coleta status do pedido.
    def ColetarStatusPedido(self) -> str:

        """
        Resumo:
        * Coleta o status do pedido no site.
        ---
        Parâmetros:
        * ===
        ---
        Retorna:
        * StatusPedido -> ["CANCELADO"] - ["FATURADO"] - ["RECUSADO"] - ["LIBERADO"] - ["RECEBIDO"].
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """

        # ~~ Coleta status do pedido.
        try: 
            StatusPedido = self.Driver.find_element(By.NAME, value="distribution_centers[1][status]")
        except: 
            try:
                StatusPedido = self.Driver.find_element(By.NAME, value="distribution_centers[2][status]")
            except:
                StatusPedido = self.Driver.find_element(By.NAME, value="distribution_centers[3][status]") 
        StatusPedido = Select(StatusPedido) 
        StatusPedido = StatusPedido.first_selected_option.text

        # ~~ Converte status.
        if StatusPedido == "Cancelado pela positivo":
            StatusPedido = "CANCELADO"
        elif StatusPedido in ["Expedido", "Expedido parcial"]:
            StatusPedido = "FATURADO"
        elif StatusPedido == "Recusado pelo crédito":
            StatusPedido = "RECUSADO"
        elif StatusPedido in ["Pedido integrado", "Em separação", "Crédito aprovado", "Faturado"]:
            StatusPedido = "LIBERADO"
        elif StatusPedido == "Pedido recebido":
            StatusPedido = "RECEBIDO"

        # ~~ Retorna status.
        return StatusPedido

    # ================================================== #

    # ~~ Coleta cliente do pedido.
    def ColetarClientePedido(self) -> str:

        """
        Resumo:
        * Coleta o cliente do pedido no site.
        ---
        Parâmetros:
        * ===
        ---
        Retorna:
        * Cliente -> Razão social
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """

        # ~~ Coleta cliente.
        Cliente = self.Driver.find_element(By.XPATH, value="//label[@for='client_name_corporate']/following-sibling::div[@class='col-md-12']").text
        Cliente = str(Cliente).split(" (")[0]

        # ~~ Retorna.
        return Cliente

    # ================================================== #

    # ~~ Abrir transação.
    def AbrirTransação(self, Transação: str) -> None:

        """
        Resumo:
        * Abre transação no SAP.
        ---
        Parâmetros:
        * Transação -> Código da transação SAP.
        ---
        Retorna:
        * ===
        ---
        Erros:
        * Sem acesso à {Transação}.
        ---
        Erros tratados localmente:
        * Sem acesso à {Transação}.
        ---
        Erros levantados:
        * ===
        """

        # ~~ Acessa transação.
        self.Session.findById("wnd[0]/tbar[0]/okcd").text = "/N" + Transação
        self.Session.findById("wnd[0]").sendVKey(0)
        StatusBarMsg = None
        StatusBarMsg = self.Session.findById("wnd[0]/sbar").text
        if "Sem autorização" in StatusBarMsg:
            Erro = f"Sem acesso à {Transação}."
            self.PrintarMensagem(Erro, "=", 30, "bot")
            self.EncerrarRPA()

    # ================================================== #

    # ~~ Verifica se data de vencimento da nota é válida.
    def VerificarSeEstáVencido(self, DataParaVerificar: str) -> str:

        """
        Resumo:
        * Verifica se data de vencimento da nota está vencida ou não.
        ---
        Parâmetros:
        * DataParaVerificar -> Data para verificar se está vencida.
        ---
        Retorna:
        * Resultado -> ["Vencido"] - ["Não vencido"]
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """

        DataVencido = datetime.strptime(DataParaVerificar, "%d/%m/%Y").date()
        DataAtual = datetime.now().date()
        if DataVencido < DataAtual:
            DiasVencidos = 0
            DataVencido = DataVencido + timedelta(days = 1)
            while DataVencido < DataAtual:
                if DataVencido.weekday() < 5:
                    DiasVencidos += 1
                DataVencido = DataVencido + timedelta(days = 1)
            if DiasVencidos >= 2:
                return "Vencido"
            else:
                return "Não vencido"
        else:
            return "Não vencido"

    # ================================================== #

    # ~~ Coleta dados financeiros do cliente.
    def ColetarDadosFinanceiros(self, RaizCnpj: str) -> dict:

        """
        Resumo:
        * Coleta dados financeiros do cliente.
        ---
        Parâmetros:
        * ContaFD33 -> Código ERP do cliente.
        * RaizCnpj -> Raiz do CNPJ.
        ---
        Retorna:
        * Dados -> Dicionário com: ["NfVencida] - ["EmAberto"] - ["Limite"] - ["Vencimento"].
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """

        # ~~ Cria dicionário novo para armazenar dados.
        Dados = {}

        # ~~ Acessa transação. Se der erro, retorna erro.
        self.AbrirTransação("FD33")
        
        # ~~ Coleta dados da FD33.
        self.Session.findById("wnd[0]").sendVKey(4)
        self.Session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006").select()
        i = 1
        while True:
            self.Session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]").text = ""
            self.Session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]").text = f"{RaizCnpj}000{i}*"
            self.Session.findById("wnd[1]/tbar[0]/btn[0]").press()
            Msg = self.Session.findById("wnd[0]/sbar").text
            if "Nenhum valor para esta seleção" in Msg:
                i += 1
            else:
                break
        self.Session.findById("wnd[1]").sendVKey(2)
        self.Session.findById("wnd[0]/usr/ctxtRF02L-KKBER").text = "1000"
        self.Session.findById("wnd[0]/usr/chkRF02L-D0210").selected = True
        self.Session.findById("wnd[0]").sendVKey(0)
        Limite = self.Session.findById("wnd[0]/usr/txtKNKK-KLIMK").text
        Limite = Limite.replace(".", "").replace(",", ".")
        Limite = float(Limite)
        LimiteStr = f"R$ {Limite:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        Vencimento = self.Session.findById("wnd[0]/usr/ctxtKNKK-NXTRV").text
        if not Vencimento == "":
            Vencimento = datetime.strptime(Vencimento, "%d.%m.%Y").date()
            VencimentoStr = datetime.strftime(Vencimento, "%d/%m/%Y")
            self.PrintarMensagem(f"Limite: {LimiteStr}", "=", 30, "bot")
            self.PrintarMensagem(f"Vencimento do limite: {VencimentoStr}", "=", 30, "bot")
        else:
            Vencimento = "-"
            self.PrintarMensagem(f"Limite: {LimiteStr}", "=", 30, "bot")
            self.PrintarMensagem(f"Vencimento do limite: {Vencimento}", "=", 30, "bot")

        # ~~ Acessa transação FBL5N, se der erro, retorna erro.
        self.AbrirTransação("FBL5N")

        # ~~ Coleta dados da FBL5N.
        self.Session.findById("wnd[0]/tbar[1]/btn[17]").press()
        self.Session.findById("wnd[1]/usr/txtENAME-LOW").text = "72776"
        self.Session.findById("wnd[1]/tbar[0]/btn[8]").press()
        self.Session.findById("wnd[0]").sendVKey(4)
        self.Session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006").select()
        self.Session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]").text = f"{RaizCnpj}*"
        self.Session.findById("wnd[1]").sendVKey(0)
        Contas = []
        for Linha in range(3, 50):
            try:
                Conta = self.Session.findById(f"wnd[1]/usr/lbl[119,{Linha}]").text
            except:
                continue
            if Conta != "":
                Contas.append(Conta)
            else:
                break
        self.Session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # ~~ Cria dicionários de cada linha e os adiciona no array.
        Tabela = []
        Empresas = ["1000", "3500"]
        for Conta in Contas:
            for Empresa in Empresas:
                self.Session.findById("wnd[0]/usr/ctxtDD_KUNNR-LOW").text = Conta
                self.Session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").text = Empresa
                self.Session.findById("wnd[0]/tbar[1]/btn[8]").press()
                Msg = self.Session.findById("wnd[0]/sbar").text
                FormaBusca = "SCROLL"
                if Msg not in ["Nenhuma partida selecionada (ver texto descritivo)", "Nenhuma conta preenche as condições de seleção"]:

                    # ~~ Verificando se forma de busca será por scroll na barra ou não.
                    for Linha in range(10, 100):
                        try:
                            Célula = self.Session.findById(f"wnd[0]/usr/lbl[0,{Linha}]").text
                            if Célula == " Cliente":
                                FormaBusca = "ESTÁTICO"
                                break
                        except:
                            continue

                    # ~~ Forma de busca ESTÁTICO:
                    if FormaBusca == "ESTÁTICO":
                        for Linha in range(10, 30):
                            TabelaDicionario = {}
                            try:
                                Situacao = self.Session.findById(f"wnd[0]/usr/lbl[6,{Linha}]").IconName
                                if Situacao != "S_LEDR":
                                    break
                                else:
                                    Situacao = "Em aberto"
                                FrmPag = self.Session.findById(f"wnd[0]/usr/lbl[39,{Linha}]").text
                                CondPag = self.Session.findById(f"wnd[0]/usr/lbl[132,{Linha}]").text
                                Conciliação = self.Session.findById(f"wnd[0]/usr/lbl[9,{Linha}]").text
                                Valor = self.Session.findById(f"wnd[0]/usr/lbl[62,{Linha}]").text
                                Valor = Valor.replace(" ", "")
                                if FrmPag in ["7", "2", "M", "G", "J", "Z", "V", "A", "P", "S", "*"] or CondPag in ["0001", "0002", "Z576", "Z577"]:
                                    if not Valor.endswith("-"):
                                        continue
                                Vencido =  self.Session.findById(f"wnd[0]/usr/lbl[42,{Linha}]").IconName
                                if Vencido == "RESUBM":
                                    Vencido = "No prazo"
                                else:
                                    Conciliação = self.Session.findById(f"wnd[0]/usr/lbl[9,{Linha}]").text
                                    Texto = self.Session.findById(f"wnd[0]/usr/lbl[81,{Linha}]").text
                                    if Conciliação == "CONCILIACAO":
                                        Vencido = "Conciliação"
                                    elif "DEVOLUÇÃO" in Texto:
                                        Vencido = "Devolução"
                                    elif "EXTRAVIO" in Texto:
                                        Vencido = "Extravio"
                                    else:
                                        try:
                                            DataVencimento = Conciliação
                                            if "." in DataVencimento:
                                                DataVencimento = str(DataVencimento).replace(".", "/")
                                            DataVencimentoDat = datetime.strptime(DataVencimento, "%d/%m/%Y")
                                        except:
                                            DataVencimento = self.Session.findById(f"wnd[0]/usr/lbl[28,{Linha}]").text
                                            DataVencimento = str(DataVencimento).replace(".", "/")
                                        Resultado = self.VerificarSeEstáVencido(DataVencimento)
                                        if Resultado == "Vencido":
                                            Vencido = "Vencido"
                                        else:
                                            Vencido = "No prazo"
                                Nf = self.Session.findById(f"wnd[0]/usr/lbl[45,{Linha}]").text
                                if Nf == "":
                                    break
                                if Valor.endswith("-"):
                                    Valor = "-" + Valor[:-1]
                                    Vencido = "Crédito"
                                TabelaDicionario["CONTA"] = Conta
                                TabelaDicionario["SITUAÇÃO"] = Situacao
                                TabelaDicionario["FRM. PAGAMENTO"] = FrmPag
                                TabelaDicionario["CND. PAGAMENTO"] = CondPag
                                TabelaDicionario["VENCIMENTO"] = Vencido
                                TabelaDicionario["NF"] = Nf
                                TabelaDicionario["VALOR"] = Valor
                                Tabela.append(TabelaDicionario)
                            except:
                                break

                    # ~~ Forma de busca SCROLL:
                    else:
                        for Linha in range(0, 500):
                            TabelaDicionario = {}
                            self.Session.findById("wnd[0]/usr").verticalScrollbar.position = Linha
                            try:
                                Situacao = self.Session.findById(f"wnd[0]/usr/lbl[6,10]").IconName
                                if Situacao != "S_LEDR":
                                    break
                                else:
                                    Situacao = "Em aberto"
                                FrmPag = self.Session.findById(f"wnd[0]/usr/lbl[39,10]").text
                                CondPag = self.Session.findById(f"wnd[0]/usr/lbl[132,10]").text
                                Conciliação = self.Session.findById(f"wnd[0]/usr/lbl[9,10]").text
                                Valor = self.Session.findById(f"wnd[0]/usr/lbl[62,10]").text
                                Valor = Valor.replace(" ", "")
                                if FrmPag in ["7", "2", "M", "G", "J", "Z", "V", "A", "P", "S", "*"] or CondPag in ["0001", "0002", "Z576", "Z577"]:
                                    if not Valor.endswith("-"):
                                        continue
                                Vencido =  self.Session.findById(f"wnd[0]/usr/lbl[42,10]").IconName
                                if Vencido == "RESUBM":
                                    Vencido = "No prazo"
                                else:
                                    Conciliação = self.Session.findById(f"wnd[0]/usr/lbl[9,10]").text
                                    Texto = self.Session.findById("wnd[0]/usr/lbl[81,10]").text
                                    if Conciliação == "CONCILIACAO":
                                        Vencido = "Conciliação"
                                    elif "DEVOLUÇÃO" in Texto:
                                        Vencido = "Devolução"
                                    elif "EXTRAVIO" in Texto:
                                        Vencido = "Extravio"
                                    else:
                                        try:
                                            DataVencimento = Conciliação
                                            if "." in DataVencimento:
                                                DataVencimento = str(DataVencimento).replace(".", "/")
                                            DataVencimentoDat = datetime.strptime(DataVencimento, "%d/%m/%Y")
                                        except:
                                            DataVencimento = self.Session.findById(f"wnd[0]/usr/lbl[28,10]").text
                                            DataVencimento = str(DataVencimento).replace(".", "/")
                                        Resultado = self.VerificarSeEstáVencido(DataVencimento)
                                        if Resultado == "Vencido":
                                            Vencido = "Vencido"
                                        else:
                                            Vencido = "No prazo"
                                Nf = self.Session.findById(f"wnd[0]/usr/lbl[45,10]").text
                                if Nf == "":
                                    break
                                if Valor.endswith("-"):
                                    Valor = "-" + Valor[:-1]
                                    Vencido = "Crédito"
                                TabelaDicionario["CONTA"] = Conta
                                TabelaDicionario["SITUAÇÃO"] = Situacao
                                TabelaDicionario["FRM. PAGAMENTO"] = FrmPag
                                TabelaDicionario["CND. PAGAMENTO"] = CondPag
                                TabelaDicionario["VENCIMENTO"] = Vencido
                                TabelaDicionario["NF"] = Nf
                                TabelaDicionario["VALOR"] = Valor
                                Tabela.append(TabelaDicionario)
                            except:
                                break
                    self.Session.findById("wnd[0]").sendVKey(3)
                else:
                    continue

        # ~~ Cria um data frame utilizando pandas. Se não consegue criar, é porque cliente não possui nada em aberto. Também coleta os vencidos caso tenha.
        if Tabela:
            Df = pd.DataFrame(Tabela)
            Df["VALOR"] = Df["VALOR"].str.replace(".", "").str.replace(",", ".")
            Df["VALOR"] = Df["VALOR"].astype(float)

            # ~~ Faz soma do total em aberto e adiciona no data frame.
            SomaTotal = Df["VALOR"].sum()
            NovaLinha = pd.DataFrame({"CONTA": [""], "SITUAÇÃO": [""], "FRM. PAGAMENTO": [""], "CND. PAGAMENTO": [""], "VENCIMENTO": [""], "NF": ["TOTAL"], "VALOR": [SomaTotal]})
            Df = pd.concat([Df, NovaLinha])
            self.PrintarMensagem(f"Valores em aberto do cliente: {RaizCnpj}\n{Df}", "=", 30, "bot")
            EmAberto = SomaTotal
            TotalLinhas = Df.shape[0]
            Mensagem = "As seguintes notas estão vencidas:\n"
            NfVencida = ""
            for Linha in range(0, TotalLinhas):
                if Df.iloc[Linha]["VENCIMENTO"] == "Vencido":
                    if NfVencida == "":
                        NfVencida += Df.iloc[Linha]["NF"]
                    else:
                        NfVencida += " || " + Df.iloc[Linha]["NF"]
            Mensagem = Mensagem + NfVencida
            if NfVencida == "":
                self.PrintarMensagem("Sem vencidos.", "=", 30, "bot")
            else:
                self.PrintarMensagem(Mensagem, "=", 30, "bot")
        else:
            NfVencida = ""
            self.PrintarMensagem(f"Cliente: {RaizCnpj} não possui nada em aberto.", "=", 30, "bot")
            EmAberto = 0

        # ~~ Adiciona tudo no dicionário.
        Dados["NfVencida"] = NfVencida
        Dados["EmAberto"] = EmAberto
        Dados["Limite"] = Limite
        Dados["Vencimento"] = Vencimento

        # ~~ Retorna.
        return Dados

    # ================================================== #

    # ~~ Coleta vendedor do pedido.
    def ColetarVendedorPedido(self) -> str:

        """
        Resumo:
        * Coleta vendedor do pedido.
        ---
        Parâmetros:
        * ===
        ---
        Retorna:
        * Vendedor -> Nome do vendedor.
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """
        
        # ~~ Coleta vendedor.
        Cnpj = self.Driver.find_element(By.XPATH, value="//label[@for='client_cnpj']/following-sibling::div[@class='col-md-12']").text 
        self.Driver.get("https://www.revendedorpositivo.com.br/admin/clients")
        Pesquisa = self.Driver.find_element(By.ID, value="keyword") 
        Pesquisa.clear()
        Pesquisa.send_keys(Cnpj)
        Pesquisa.send_keys(Keys.ENTER)
        time.sleep(3)
        try:
            Editar = self.Driver.find_elements(By.XPATH, value="//table/tbody/tr")[1].find_elements(By.XPATH, value=".//td")[10].find_element(By.XPATH, value=".//a") 
            Editar = Editar.get_attribute("href")
            self.Driver.get(str(Editar)) 
            time.sleep(3)
            Carteira = self.Driver.find_element(By.XPATH, value="//section").find_elements(By.XPATH, value=".//ul/li")[10].find_element(By.XPATH, value=".//a")
            Carteira.click()
            Carteira = self.Driver.find_element(By.XPATH, value="(//select[@class='form-control select-multiple side2side-selected-options side2side-select-taller'])[1]")
            Carteira = Select(Carteira)
            Carteira = Carteira.options
            Vendedor = Carteira[0].text
        except:

            # ~~ Se não encontra vendedor 1105, tenta encontrar pelo 1101 nos ativos.
            try:
                self.Driver.get("https://www.revendedorpositivo.com.br/admin/direct-billing-clients")
                Pesquisa = self.Driver.find_element(By.ID, value="keyword") 
                Pesquisa.clear() 
                Pesquisa.send_keys(Cnpj) 
                Ativo = self.Driver.find_element(By.ID, value="active-1")
                Ativo.click()
                Pesquisa.send_keys(Keys.ENTER)
                time.sleep(3)
                Editar = self.Driver.find_elements(By.XPATH, value="//table/tbody/tr")[1].find_element(By.XPATH, value="//td[contains(@data-title, 'Ações')]/a").get_attribute("href")
                self.Driver.get(str(Editar))
                Cnpj = self.Driver.find_element(By.ID, value="resale_cnpj").get_attribute("value")
                self.Driver.get("https://www.revendedorpositivo.com.br/admin/clients")
                Pesquisa = self.Driver.find_element(By.ID, value="keyword")
                Pesquisa.clear() 
                Pesquisa.send_keys(Cnpj) 
                Pesquisa.send_keys(Keys.ENTER)
                time.sleep(3)
                Editar = self.Driver.find_elements(By.XPATH, value="//table/tbody/tr")[1].find_elements(By.XPATH, value=".//td")[10].find_element(By.XPATH, value=".//a") 
                Editar = Editar.get_attribute("href") 
                self.Driver.get(str(Editar)) 
                time.sleep(3)
                Carteira = self.Driver.find_element(By.XPATH, value="//section").find_elements(By.XPATH, value=".//ul/li")[10].find_element(By.XPATH, value=".//a") 
                Carteira.click() 
                Carteira = self.Driver.find_element(By.XPATH, value="(//select[@class='form-control select-multiple side2side-selected-options side2side-select-taller'])[1]") 
                Carteira = Select(Carteira) 
                Carteira = Carteira.options 
                Vendedor = Carteira[0].text 
            
            # ~~ Se não encontrar 1101 ativos, procura nos inativos.
            except:
                self.Driver.get("https://www.revendedorpositivo.com.br/admin/direct-billing-clients")
                Pesquisa = self.Driver.find_element(By.ID, value="keyword") 
                Pesquisa.clear() 
                Pesquisa.send_keys(Cnpj) 
                Inativo = self.Driver.find_element(By.ID, value="active-0")
                Inativo.click()
                Pesquisa.send_keys(Keys.ENTER)
                time.sleep(3)
                Editar = self.Driver.find_elements(By.XPATH, value="//table/tbody/tr")[1].find_element(By.XPATH, value="//td[contains(@data-title, 'Ações')]/a").get_attribute("href")
                self.Driver.get(str(Editar))
                Cnpj = self.Driver.find_element(By.ID, value="resale_cnpj").get_attribute("value")
                self.Driver.get("https://www.revendedorpositivo.com.br/admin/clients")
                Pesquisa = self.Driver.find_element(By.ID, value="keyword")
                Pesquisa.clear() 
                Pesquisa.send_keys(Cnpj) 
                Pesquisa.send_keys(Keys.ENTER)
                time.sleep(3)
                Editar = self.Driver.find_elements(By.XPATH, value="//table/tbody/tr")[1].find_elements(By.XPATH, value=".//td")[10].find_element(By.XPATH, value=".//a") 
                Editar = Editar.get_attribute("href") 
                self.Driver.get(str(Editar)) 
                time.sleep(3)
                Carteira = self.Driver.find_element(By.XPATH, value="//section").find_elements(By.XPATH, value=".//ul/li")[10].find_element(By.XPATH, value=".//a") 
                Carteira.click() 
                Carteira = self.Driver.find_element(By.XPATH, value="(//select[@class='form-control select-multiple side2side-selected-options side2side-select-taller'])[1]") 
                Carteira = Select(Carteira)
                Carteira = Carteira.options
                Vendedor = Carteira[0].text
        
        # ~~ Retorna.
        return Vendedor

    # ================================================== #

    # ~~ Coleta dados do pedido no site em recursividade até encontrar forma de crédito interno.
    def ColetarDadosPedido(self, Pedido: int) -> dict:

        """
        Resumo:
        * Coleta dados do pedido no site em recursividade até encontrar forma de crédito interno.
        ---
        Parâmetros:
        * Pedido -> Nº do pedido.
        ---
        Retorna:
        * DadosPedido -> Dicionário contendo: ["Data"] - ["CondiçãoPagamento"] - ["CNPJCliente"] - ["CódigoERP"] - ["ValorPedido"] - ["Status"]
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """

        # ~~ Cria loop while true até encontrar pedido que seja crédito interno.
        while True:

            # ~~ Cria dicionário para os dados do pedido.
            DadosPedido = {}

            # ~~ Acessa página no site.
            self.AcessarPedido(Pedido)

            # ~~ Printa mensagem.
            self.PrintarMensagem(f"Coletando dados do pedido {Pedido}.", "=", 30, "bot")

            # ~~ Coleta conteúdo da página.
            ConteúdoPágina = self.Driver.find_element(By.TAG_NAME, value="body").text

            # ~~ Se pedido não foi inputado ainda, define variável para resetar loop.
            if "Application error: Mysqli statement execute error" in ConteúdoPágina:
                self.PrintarMensagem(f"Pedido {Pedido} não inserido no site ainda.", "=", 30, "bot")
                self.ReiniciarLoop = True
                return DadosPedido

            # ~~ Se pedido foi inputado, coleta forma de pagamento.
            FormaPagamento = self.ColetarFormaPagamentoPedido()

            # ~~ Se forma de pagamento for crédito interno, coleta dados do pedido e retorna.
            if FormaPagamento == "Boleto a Prazo":
                DadosPedido["Pedido"] = Pedido
                DadosPedido["Data"] = self.ColetarDataPedido()
                DadosPedido["CondiçãoPagamento"] = self.ColetarCondiçãoPagamento()
                DadosPedido["Razão"] = self.ColetarClientePedido()
                DadosPedido["CNPJCliente"] = self.ColetarCnpj()
                DadosPedido["CódigoERP"] = self.ColetarCódigoERP(DadosPedido["CNPJCliente"])
                DadosPedido["ValorPedido"] = self.ColetarValorPedido()
                DadosPedido["Status"] = self.ColetarStatusPedido()
                DadosPedido["Vendedor"] = self.ColetarVendedorPedido()
                return DadosPedido
            
            # ~~ Se não for crédito interno, usa recursão para verificar próximo pedido.
            self.PrintarMensagem(f"Pedido {Pedido} não possui forma de pagamento como crédito interno.", "=", 30, "bot")
            Pedido += 1

    # ================================================== #

    # ~~ Remove valor de liberação do controle.
    def RemoverValorLiberadoDoControle(self, Pedido: int, AdicionarEmAberto: bool) -> None:

        """
        Resumo:
        * Remove valor de liberação do pedido do controle.
        ---
        Parâmetros:
        * Pedido -> Número do Pedido.
        * AdicionarEmAberto -> True ou False para informar se é para pegar o valor removido e adicionar ao valor em aberto.
        ---
        Retorna:
        * ===
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """
        
        # ~~ Remove valor liberado.
        Path = self.Controle["BOOK"].fullname
        Df = pd.read_excel(Path, "LIMITES")
        Colunas = [
                    "PEDIDO 1", "PEDIDO 2", "PEDIDO 3", "PEDIDO 4", "PEDIDO 5", "PEDIDO 6", "PEDIDO 7", "PEDIDO 8", "PEDIDO 9", "PEDIDO 10",
                    "PEDIDO 11", "PEDIDO 12", "PEDIDO 13", "PEDIDO 14", "PEDIDO 15", "PEDIDO 16", "PEDIDO 17", "PEDIDO 18", "PEDIDO 19", "PEDIDO 20"
                    ]
        ColunasPedido = ["F", "H", "J", "L", "N", "P", "R", "T", "V", "X", "Z", "AB", "AD", "AF", "AH", "AJ", "AL", "AN", "AP", "AR"]
        ColunasValor = ["G", "I", "K", "M", "O", "Q", "S", "U", "W", "Y", "AA", "AC", "AE", "AG", "AI", "AK", "AM", "AO", "AQ", "AS"]
        for i in range(0, 20):
            Linha = Df.index[Df[Colunas[i]] == Pedido].tolist()
            if Linha:
                ColunaPedido = ColunasPedido[i]
                ColunaValor = ColunasValor[i]
                Linha = int(Linha[0])
                Linha = Linha + 2
                break
        if not Linha:
            self.PrintarMensagem("Pedido não encontrado no controle das liberações.", "=", 30, "bot")
            return
        ValorPedido = float(self.Controle["LIMITES"].range(ColunaValor + str(Linha)).value)
        ValorAberto = float(self.Controle["LIMITES"].range("D" + str(Linha)).value)
        Soma = ValorPedido + ValorAberto
        if AdicionarEmAberto == True:
            self.Controle["LIMITES"].range("D" + str(Linha)).value = Soma
            self.PrintarMensagem(f"Pedido: {Pedido} removido das liberações e somado seu valor com o que cliente tem em aberto.", "=", 30, "bot")
        else:
            self.PrintarMensagem(f"Pedido: {Pedido} removido das liberações, pois não foi faturado.", "=", 30, "bot")
        self.Controle["LIMITES"].range(ColunaPedido + str(Linha)).value = ""
        self.Controle["LIMITES"].range(ColunaValor + str(Linha)).value = ""

    # ================================================== #

    # ~~ Salva controle.
    def SalvarControle(self) -> None:

        """
        Resumo:
        * Salva o controle.
        ---
        Parâmetros:
        * ===
        ---
        Retorna:
        * ===
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """

        # ~~ Tenta salvar até 10x.
        Tentativa = 0
        while Tentativa != 10:
            try:
                self.Controle["BOOK"].save()
                return
            except:
                Tentativa += 1
                time.sleep(2)

    # ================================================== #

    # ~~ Coleta última linha preenchida na coluna.
    def ÚltimaLinhaPreenchida(self, Aba: str, Coluna: str) -> int:

        """
        Resumo:
        * Retorna última linha preenchida na coluna.
        ---
        Parâmetros:
        * Aba -> Nome da aba.
        * Coluna -> Coluna.
        ---
        Retorna:
        * ÚltimaLinha -> Número da última linha preenchida.
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """

        # ~~ Coleta última linha.
        ÚltimaLinha = self.Controle[Aba].range(Coluna + str("99999")).end("up").row

        # ~~ Retorna.
        return ÚltimaLinha

    # ================================================== #

    # ~~ Importa dados financeiros do cliente no controle.
    def ImportarDadosFinanceirosNoControle(self, Cliente: int, Vencimento: datetime = None, 
                                        Limite: float = None, EmAberto: float = None,
                                        Pedido: int = None, ValorPedido: float = None) -> None:

        """
        Resumo:
        * Importa dados financeiros do cliente no controle.
        ---
        Parâmetros:
        * Cliente -> Código ERP do cliente.
        * Vencimento (opcional) -> Data do vencimento.
        * Limite (opcional) -> Valor do limite.
        * EmAberto (opcional) -> Valor em aberto.
        * Pedido (opcional) -> Número do pedido.
        * ValorPedido (opcional) -> Valor do pedido.
        ---
        Retorna:
        * ===
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """

        # ~~ Inicia importação.
        self.SalvarControle()
        Path = self.Controle["BOOK"].fullname
        Df = pd.read_excel(Path, sheet_name = "LIMITES", dtype = {"CLIENTE": str})
        Linha = Df.index[Df['CLIENTE'] == str(Cliente)].tolist()
        if Linha:
            Linha = int(Linha[0])
            Linha = Linha + 2
        else:
            Linha = self.ÚltimaLinhaPreenchida("LIMITES", "A")
            Linha = Linha + 1
            self.Controle["LIMITES"].range("A" + str(Linha)).value = Cliente
        if Vencimento:
            self.Controle["LIMITES"].range("B" + str(Linha)).value = Vencimento
        if Limite or Limite == 0.0:
            self.Controle["LIMITES"].range("C" + str(Linha)).value = Limite
        if EmAberto or EmAberto == 0:
            self.Controle["LIMITES"].range("D" + str(Linha)).value = EmAberto
        if Pedido:
            ColunasPedidos = ["F", "H", "J", "L", "N", "P", "R", "T", "V", "X", "Z", "AB", "AD", "AF", "AH", "AJ", "AL", "AN", "AP", "AR"]
            ColunasValores = ["G", "I", "K", "M", "O", "Q", "S", "U", "W", "Y", "AA", "AC", "AE", "AG", "AI", "AK", "AM", "AO", "AQ", "AS"]
            for i in range(0, 20):
                Célula = self.Controle["LIMITES"].range(ColunasPedidos[i] + str(Linha)).value
                if Célula == Pedido:
                    self.Controle["LIMITES"].range(ColunasPedidos[i] + str(Linha)).value = Pedido
                    self.Controle["LIMITES"].range(ColunasValores[i] + str(Linha)).value = ValorPedido
                    break
                if Célula is None:
                    self.Controle["LIMITES"].range(ColunasPedidos[i] + str(Linha)).value = Pedido
                    self.Controle["LIMITES"].range(ColunasValores[i] + str(Linha)).value = ValorPedido
                    break

    # ================================================== #

    # ~~ Coleta margem do cliente.
    def ColetarMargem(self, Cliente: str) -> float:

        """
        Resumo:
        * Coleta margem do cliente que consta no banco de dados.
        ---
        Parâmetros:
        * Cliente -> Código ERP do cliente.
        ---
        Retorna:
        * Margem -> Valor da margem disponível do cliente.
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """
        
        # ~~ Coleta margem.
        self.SalvarControle()
        Path = self.Controle["BOOK"].fullname
        Df = pd.read_excel(Path, sheet_name = "LIMITES", dtype = {"CLIENTE": str})
        Linha = Df.index[Df['CLIENTE'] == str(Cliente)].tolist()
        Linha = int(Linha[0])
        Linha = Linha + 2
        Margem = float(self.Controle["LIMITES"].range("E" + str(Linha)).value)

        # ~~ Retorna.
        return Margem

    # ================================================== #

    # ~~ Faz análise de crédito do pedido.
    def AnáliseCréditoPedido(self, Pedido: int, RaizCnpj: int, Valor: float) -> dict:

        """
        Resumo:
        * Faz análise de crédito do pedido.
        ---
        Parâmetros:
        * Pedido -> Nº do pedido.
        * RaizCnpj -> Raiz do CNPJ.
        * Valor -> Valor do pedido.
        ---
        Retorna:
        * RespostaAnálise -> Dicionário contendo: ["MENSAGEM"] - ["STATUS"].
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """

        # ~~ Cria dicionário para os dados da análise.
        RespostaAnálise = {}

        # ~~ Coleta dados financeiros do cliente.
        DadosFinanceiros = self.ColetarDadosFinanceiros(RaizCnpj = RaizCnpj)

        # ~~ Organiza dados para análise.
        Cliente = str(RaizCnpj)
        DataAtual = datetime.now().date()
        Vencimento = DadosFinanceiros["Vencimento"] 
        Limite = DadosFinanceiros["Limite"]
        NfVencida = DadosFinanceiros["NfVencida"]
        EmAberto = DadosFinanceiros["EmAberto"]
        ValorPedido = Valor
        ValorPedidoStr = f"R$ {ValorPedido:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        # ~~ Printa valor do pedido.
        self.PrintarMensagem(f"Valor do pedido: {ValorPedidoStr}", "=", 30, "bot")

        # ~~ Converte Vencimento de datetime para string.
        if Vencimento == "-":
            VencimentoStr = "-"
        else:
            VencimentoStr = datetime.strftime(Vencimento, "%d/%m/%Y")

        # ~~ Importa dados financeiros do cliente no controle.
        self.ImportarDadosFinanceirosNoControle(Cliente, Vencimento, Limite, EmAberto)

        # ~~ Verifica margem.
        Margem = self.ColetarMargem(Cliente)
        Margem = round(Margem, 2)
        MargemStr = f"R$ {Margem:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        self.PrintarMensagem(f"Margem: {MargemStr}", "=", 30, "bot")

        # ~~ Inicia análise.
        LimiteAtivo = True
        Motivos = ""
        Status = "LIBERADO"

        # ~~ Verifica se possui limite ativo.
        if Limite == 0.0 or Vencimento == "-":
            Motivos += "\n- Sem limite de crédito ativo."
            Status = "NÃO LIBERADO"
            LimiteAtivo = False

        # ~~ Verifica vencimento do limite.
        elif Vencimento < DataAtual:
            Motivos += f"\n- Limite vencido em {VencimentoStr}."
            Status = "NÃO LIBERADO"
            LimiteAtivo = False

        # ~~ Verifica se pedido está dentro da margem.
        if LimiteAtivo == True:
            if Margem < ValorPedido:
                Motivos += f"\n- Valor do pedido excede a margem disponível. Valor do pedido: {ValorPedidoStr} / Margem livre: {MargemStr}."
                Status = "NÃO LIBERADO"

        # ~~ Verifica se possui notas vencidas.
        if NfVencida != "":
            Motivos += f"\n- Possui vencidos: {NfVencida}."
            Status = "NÃO LIBERADO"

        # ~~ Verifica se pode ser liberado. Se puder, importa seus dados no controle.
        if Status == "LIBERADO":
            RespostaAnálise["MENSAGEM"] = f"Pedido {Pedido} liberado."
            RespostaAnálise["STATUS"] = "LIBERADO"
            self.ImportarDadosFinanceirosNoControle(Cliente = Cliente, Pedido = Pedido, ValorPedido = ValorPedido)
        else:
            RespostaAnálise["MENSAGEM"] = f"Pedido {Pedido} recusado:{Motivos}"
            RespostaAnálise["STATUS"] = "NÃO LIBERADO"

        # ~~ Retorna com dados de liberação.
        self.PrintarMensagem(RespostaAnálise["MENSAGEM"], "=", 30, "bot")
        return RespostaAnálise

    # ================================================== #

    # ~~ Função que fica em loop coletando pedidos e fazendo análises de crédito.
    def Loop(self) -> None:

        """
        Resumo:
        * Função que fica em loop coletando pedidos e fazendo análises de crédito.
        ---
        Parâmetros:
        * ===
        ---
        Retorna:
        * ===
        ---
        Erros:
        * Exception
        ---
        Erros tratados localmente:
        * Exception
        ---
        Erros levantados:
        * ===
        """

        # ~~ Loop while true, para manter automação funcionando.
        while True:

            # ~~ Bloco try-except para automação se auto reiniciar em caso de erros.
            try:

                # ~~ Cria loop para verificar linha a linha.
                for Linha in range(2, 999999):

                    # ~~ Verifica se foi alterado o encerramento global.
                    if self.Encerrar == True:
                        self.EncerrarRPA()

                    # ~~ Coleta nº do pedido.
                    Pedido = self.Controle["PEDIDOS"].range("A" + str(Linha)).value
                    PrimeiroPedido = self.Controle["PEDIDOS"].range("B2").value

                    # ~~ Se pedido for nulo, verifica se há novas inputações no site.
                    if Pedido is None or PrimeiroPedido is None:

                        # ~~ Verifica se existe um primeiro pedido para verificar. Se existe, pega último pedido no controle + 1.
                        if PrimeiroPedido is not None:
                            Pedido = int(self.Controle["PEDIDOS"].range("A" + str(Linha - 1)).value)
                            Pedido = Pedido + 1
                        else:
                            Pedido = int(Pedido)

                        # ~~ Coleta dados do pedido no site.
                        DadosPedido = self.ColetarDadosPedido(Pedido)

                        # ~~ Se não há inputação nova de pedido, reinicia loop.
                        if self.ReiniciarLoop == True:
                            self.ReiniciarLoop = False
                            break

                        # ~~ Insere dados no controle.
                        self.PrintarMensagem(f"Inserindo dados do pedido {DadosPedido["Pedido"]} no controle.", "=", 30, "bot")
                        self.Controle["PEDIDOS"].range("A" + str(Linha)).value = DadosPedido["Pedido"]
                        self.Controle["PEDIDOS"].range("B" + str(Linha)).value = DadosPedido["Data"]
                        self.Controle["PEDIDOS"].range("C" + str(Linha)).value = DadosPedido["CondiçãoPagamento"]
                        self.Controle["PEDIDOS"].range("D" + str(Linha)).value = DadosPedido["Vendedor"]
                        self.Controle["PEDIDOS"].range("E" + str(Linha)).value = DadosPedido["Razão"]
                        self.Controle["PEDIDOS"].range("F" + str(Linha)).value = DadosPedido["CNPJCliente"]
                        self.Controle["PEDIDOS"].range("G" + str(Linha)).value = DadosPedido["ValorPedido"]
                        self.Controle["PEDIDOS"].range("H" + str(Linha)).value = DadosPedido["Status"]
                        self.Controle["PEDIDOS"].range("I" + str(Linha)).value = "-"
                        self.Controle["PEDIDOS"].range("J" + str(Linha)).value = "-"
                        self.Controle["PEDIDOS"].range("K" + str(Linha)).value = "-"

                    # ~~ Coleta dados do pedido relevantes para análise de crédito.
                    Pedido = int(self.Controle["PEDIDOS"].range("A" + str(Linha)).value)
                    DataPedido = self.Controle["PEDIDOS"].range("B" + str(Linha)).value
                    Vendedor = str(self.Controle["PEDIDOS"].range("D" + str(Linha)).value)
                    RaizCnpj = str(self.Controle["PEDIDOS"].range("F" + str(Linha)).value)
                    ValorPedido = float(self.Controle["PEDIDOS"].range("G" + str(Linha)).value)
                    Status = str(self.Controle["PEDIDOS"].range("H" + str(Linha)).value)
                    Email = str(self.Controle["PEDIDOS"].range("K" + str(Linha)).value)

                    # ~~ Se status for "LIBERADO".
                    if Status == "LIBERADO":

                        # ~~ Verifica se houve atualização de status para "FATURADO" ou "CANCELADO".
                        self.PrintarMensagem(f"Verificando se houve atualização de status do pedido {Pedido}.", "=", 30, "bot")
                        self.AcessarPedido(Pedido)
                        StatusNovo = self.ColetarStatusPedido()

                        # ~~ Se status novo for "FATURADO", adiciona valor de liberação ao em aberto do cliente.
                        if StatusNovo == "FATURADO":
                            self.PrintarMensagem(f"Status do {Pedido} atualizado: {Status} => {StatusNovo}.", "=", 30, "bot")
                            self.Controle["PEDIDOS"].range("H" + str(Linha)).value = StatusNovo
                            self.RemoverValorLiberadoDoControle(Pedido = Pedido, AdicionarEmAberto = True)

                        # ~~ Se status novo for "CANCELADO", remove valor de liberação do pedido.
                        elif StatusNovo == "CANCELADO":
                            self.PrintarMensagem(f"Status do {Pedido} atualizado: {Status} => {StatusNovo}.", "=", 30, "bot")
                            self.Controle["PEDIDOS"].range("H" + str(Linha)).value = StatusNovo
                            self.RemoverValorLiberadoDoControle(Pedido = Pedido, AdicionarEmAberto = False)

                        # ~~ Se não houve atualização de status.
                        else:
                            self.PrintarMensagem(f"Pedido {Pedido} sem atualização de status.", "=", 30, "bot")

                    # ~~ Se status for "RECEBIDO".
                    if Status == "RECEBIDO":

                        # ~~ Verifica se houve atualização de status para "CANCELADO".
                        self.PrintarMensagem(f"Verificando se houve atualização de status do pedido {Pedido}.", "=", 30, "bot")
                        self.AcessarPedido(Pedido)
                        StatusNovo = self.ColetarStatusPedido()

                        # ~~ Se status tiver atualizado para "CANCELADO", atualiza controle.
                        if StatusNovo == "CANCELADO":
                            self.PrintarMensagem(f"Status do {Pedido} atualizado para {StatusNovo}.", "=", 30, "bot")
                            self.Controle["PEDIDOS"].range("H" + str(Linha)).value = StatusNovo

                        # ~~ Se não houve atualização de status, faz análise.
                        else:
                            self.PrintarMensagem(f"Iniciando análise do pedido {Pedido}. Possui status {Status}.", "=", 30, "bot")
                            RespostaAnálise = self.AnáliseCréditoPedido(Pedido = Pedido, RaizCnpj = RaizCnpj, Valor = ValorPedido)
                            self.Controle["PEDIDOS"].range("I" + str(Linha)).value = RespostaAnálise["MENSAGEM"]
                            self.Controle["PEDIDOS"].range("J" + str(Linha)).value = datetime.now()

                            # ~~ Verifica resposta de liberação e atualiza status no controle e no site.
                            if RespostaAnálise["STATUS"] == "NÃO LIBERADO":
                                self.AlterarPedidoSite(Pedido = Pedido, AlterarStatus = "Recusado pelo crédito", ObservaçãoInterna = RespostaAnálise["MENSAGEM"])
                                self.Controle["PEDIDOS"].range("H" + str(Linha)).value = "RECUSADO"
                            else:
                                self.AlterarPedidoSite(Pedido = Pedido, AlterarStatus = "Crédito aprovado", ObservaçãoInterna = RespostaAnálise["MENSAGEM"])
                                self.Controle["PEDIDOS"].range("H" + str(Linha)).value = "LIBERADO"

                    # ~~ Se status for "RECUSADO".
                    if Status == "RECUSADO":

                        # ~~ Verifica se houve atualização de status para "CANCELADO".
                        self.PrintarMensagem(f"Verificando se houve atualização de status do pedido {Pedido}.", "=", 30, "bot")
                        self.AcessarPedido(Pedido)
                        StatusNovo = self.ColetarStatusPedido()

                        # ~~ Se status tiver atualizado para "CANCELADO", atualiza controle.
                        if StatusNovo == "CANCELADO":
                            self.PrintarMensagem(f"Status do {Pedido} atualizado para {StatusNovo}.", "=", 30, "bot")
                            self.Controle["PEDIDOS"].range("H" + str(Linha)).value = StatusNovo

                        # ~~ Se não houve atualização de status, faz reanálise.
                        else:
                            self.PrintarMensagem(f"Iniciando reanálise do pedido {Pedido}. Status continua como {Status}.", "=", 30, "bot")
                            RespostaAnálise = self.AnáliseCréditoPedido(Pedido = Pedido, RaizCnpj = RaizCnpj, Valor = ValorPedido)
                            self.Controle["PEDIDOS"].range("I" + str(Linha)).value = RespostaAnálise["MENSAGEM"]
                            self.Controle["PEDIDOS"].range("J" + str(Linha)).value = datetime.now()

                            # ~~ Verifica resposta de liberação e atualiza status no controle e no site.
                            if RespostaAnálise["STATUS"] == "NÃO LIBERADO":
                                self.AlterarPedidoSite(Pedido = Pedido, ObservaçãoInterna = RespostaAnálise["MENSAGEM"])
                                self.Controle["PEDIDOS"].range("H" + str(Linha)).value = "RECUSADO"
                            else:
                                self.AlterarPedidoSite(Pedido = Pedido, AlterarStatus = "Crédito aprovado", ObservaçãoInterna = RespostaAnálise["MENSAGEM"])
                                self.Controle["PEDIDOS"].range("H" + str(Linha)).value = "LIBERADO"

            # ~~ Se der erro, printa ele, espera 1m e reinicia loop.
            except Exception as Erro:
                self.PrintarMensagem(f"Ocorreu o seguinte erro na execução: {Erro}. Aguardando 1m para reinício.", "=", 30, "bot")
                time.sleep(60)

    # ================================================== #

    # ~~ Encerramento do RPA.
    def EncerrarRPA(self) -> None:

        """
        Resumo:
        * Encerra RPA.
        ---
        Parâmetros:
        * ===
        ---
        Retorna:
        * ===
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """

        # ~~ Encerra instância Driver.
        self.Driver.quit()

        # ~~ Encerra instância SAP.
        while True:
            if self.Session.ActiveWindow.Text == "SAP Easy Access":
                break
            else:
                self.Session.findById("wnd[0]").sendVKey(3)
        self.Session = None
        
        # ~~ Salva controle.
        self.SalvarControle()

        # ~~ Encerra execução do RPA.
        self.PrintarMensagem("Encerrando execução do RPA...", "=", 30, "bot")
        self.ExportarLog()
        exit()

    # ================================================== #

    # ~~ Exporta log.
    def ExportarLog(self) -> None:

        """
        Resumo:
        * Exporta log do terminal em ".txt".
        ---
        Parâmetros:
        * ===
        ---
        Retorna:
        * ===
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """

        # ~~ Pega data e hora atual para nomear arquivo.
        DataHoraFim = datetime.now().replace(microsecond = 0)
        DataHoraFim = DataHoraFim.strftime("%d-%m-%Y_%H-%M")

        # ~~ Path da pasta logs.
        CaminhoScript = os.path.abspath(__file__)
        CaminhoLogs = CaminhoScript.split(r"script_crédito.py")[0] + r"\logs"

        # ~~ Exporta Log.
        with open(fr"{CaminhoLogs}\{self.DataHoraInício} & {DataHoraFim}.txt", "w", encoding = "utf-8") as LogFile:
            LogFile.write(self.Log)

    # ================================================== #

    # ~~ Printa mensagem ASCII.
    def ASCII(self) -> None:

        """
        Resumo:
        * Printa mensagem ASCII inicial.
        ---
        Parâmetros:
        * ===
        ---
        Retorna:
        * ===
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """

        # ~~ Mensagem ASCII.
        Ascii1 =  r"""#########################################################"""
        Ascii2 =  r"""#                                                       #"""
        Ascii3 =  r"""#  ____  ____   _       ____       __     _ _ _         #"""
        Ascii4 =  r"""# |  _ \|  _ \ / \     / ___|_ __ /_/  __| (_) |_ ___   #"""
        Ascii5 =  r"""# | |_) | |_) / _ \   | |   | '__/ _ \/ _` | | __/ _ \  #"""
        Ascii6 =  r"""# |  _ <|  __/ ___ \  | |___| | |  __/ (_| | | || (_) | #"""
        Ascii7 =  r"""# |_| \_\_| /_/   \_\  \____|_|  \___|\__,_|_|\__\___/  #"""
        Ascii8 =  r"""#                                                       #"""
        Ascii9 =  r"""#########################################################"""
        Ascii = f"{Ascii1}\n{Ascii2}\n{Ascii3}\n{Ascii4}\n{Ascii5}\n{Ascii6}\n{Ascii7}\n{Ascii8}\n{Ascii9}"

        # ~~ Printa.
        self.PrintarMensagem(Ascii, "=", 30, "bot")

    # ================================================== #

    # ~~ Inicia RPA.
    def IniciarRPA(self) -> None:

        """
        Resumo:
        * Inicia RPA.
        ---
        Parâmetros:
        * ===
        ---
        Retorna:
        * ===
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """

        # ~~ Printa ASCII.
        self.ASCII()

        # ~~ Cria instâncias.
        self.Session = self.InstanciarSap()
        self.Driver = self.InstanciarNavegador()
        self.Controle = self.InstanciarControle()

        # ~~ Inicia loop do RPA.
        self.Loop()

    # ================================================== #

    # ~~ Monitora o encerramento do RPA.
    def MonitarEncerramento(self) -> None:

        """
        Resumo:
        * Monitora o encerramento do RPA ao pressionar CTRL + F12.
        ---
        Parâmetros:
        * ===
        ---
        Retorna:
        * ===
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """ 

        # ~~ Mantém o script em execução, aguardando o pressionamento da tecla.
        while not self.Encerrar == True:

            # ~~ Time sleep para reduzir uso de CPU.
            time.sleep(0.5)

            # ~~ Se pressionado "CTRL+F12", inicia encerramento.
            if keyboard.is_pressed("CTRL+F12"):
                self.DefinirEncerramento()

    # ================================================== #

    # ~~ Altera variável de encerramento global.
    def DefinirEncerramento(self) -> None:

        """
        Resumo:
        * Altera variável de encerramento global.
        ---
        Parâmetros:
        * ===
        ---
        Retorna:
        * ===
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """ 

        # ~~ Define encerramento global para True.
        self.Encerrar = True

    # ================================================== #

    # ~~ Altera pedido no site.
    def AlterarPedidoSite(self, Pedido: int, AlterarStatus: str = None, ObservaçãoInterna: str = None) -> None:

        """
        Resumo:
        * Altera pedido no site. Podendo alterar o status e inserir observação interna.
        ---
        Parâmetros:
        * Pedido -> Nº do pedido.
        * AlterarStatus -> Se passado algum status, altera para ele no site.
        * ObservaçãoInterna -> Se passado alguma observação, insere ela nas observações do pedido.
        ---
        Retorna:
        * ===
        ---
        Erros:
        * ===
        ---
        Erros tratados localmente:
        * ===
        ---
        Erros levantados:
        * ===
        """ 

        # ~~ Acessa página.
        self.AcessarPedido(Pedido)

        # ~~ Se for para alterar o status.
        if AlterarStatus is not None:

            # ~~ Encontra paineis e altera status em cada um.
            for i in range(1, 4):
                try: 
                    StatusPedido = self.Driver.find_element(By.NAME, value = f"distribution_centers[{i}][status]")
                    StatusPedido = Select(StatusPedido)
                    StatusPedido.select_by_visible_text(AlterarStatus)
                except:
                    continue
        
        # ~~ Se for para inserir observação interna.
        if ObservaçãoInterna is not None:
            CampoObservação = self.Driver.find_element(By.ID, value = "comment")
            CampoObservação.clear()
            CampoObservação.send_keys(ObservaçãoInterna)

        # ~~ Salva.
        botãoSalvar = self.Driver.find_element(By.ID, value="save")
        botãoSalvar.click()

# ================================================== #

# ~~ Inicia código.
if __name__ == "__main__":

    # ~~ Cria instância do RPA.
    Rpa = RPACrédito()

    # ~~ Inicializa threading para encerramento do RPA.
    thread = threading.Thread(target = Rpa.MonitarEncerramento)
    thread.daemon = True
    thread.start()

    # ~~ Inicia RPA.
    Rpa.IniciarRPA()

# ================================================== #