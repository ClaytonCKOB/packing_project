import pyodbc
from classes.produto import *
import constants as const
class ConnDatabase:
    def __init__(self):
        #Conexão com a tabela do Sysop
        self.conn = pyodbc.connect(r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
                                   f"DBQ={const.DB_PATH};")
        self.cursor = self.conn.cursor()

        #Comandos SQL
        self.sql_cortInfo   = "SELECT cliente, ambiente, modelo, colecao, cor, Text18, acionamento, t_lado, obs, cor_trilho, alturacomando FROM t_os_desc_cortina WHERE t_os_desc_cortina.IDitem = "
        self.sql_cliente    = "SELECT cliente FROM t_os_desc_cortina WHERE t_os_desc_cortina.codos = "
        self.sql_pedido     = "SELECT codos FROM t_os_desc WHERE t_os_desc.IDItem = "
        self.sql_idprod     = "SELECT IDPro FROM t_os_desc WHERE t_os_desc.IDItem = "
        self.sql_quant      = "SELECT quant FROM t_os_desc WHERE t_os_desc.IDItem = "
        self.sql_desc       = "SELECT xProd FROM t_a_pro WHERE t_a_pro.IDPro = "
        self.sql_iditem     = "SELECT DISTINCT IDItem FROM t_os_desc WHERE t_os_desc.codos = "
        self.sql_largura    = "SELECT largura FROM t_os_desc_med WHERE "
        self.sql_altura     = "SELECT altura FROM t_os_desc_med WHERE "
        self.sql_item       = "SELECT item FROM t_os_desc WHERE "
        self.sql_IdGuia     = "SELECT IDItem, IDAce FROM t_os_desc_acess WHERE t_os_desc_acess.IDPed = "
        self.sql_txtGuia    = "SELECT texto FROM t_a_combo WHERE t_a_combo.id = "
        self.sql_unid       = "SELECT uCom FROM t_a_pro WHERE t_a_pro.IDPro = "
        self.sql_obs        = "SELECT obs FROM t_os_desc_cortina WHERE t_os_desc_cortina.IDitem = "

        self.sql_idCliente          = "SELECT IDCad FROM t_os WHERE t_os.codos = "
        self.sql_infoCliente        = "SELECT CNPJ, CEP, endereco, Bairro, Cidade, UF FROM t_cliente WHERE t_cliente.IDCad = "

        self.sql_vlrPedido          = "SELECT vlr_total FROM t_os WHERE t_os.codos = "

        self.sql_obsPedido          = "SELECT Obs FROM t_os_obs WHERE t_os_obs.codos = "

    def obterCortInfo(self, iditem):
        data = self.cursor.execute(self.sql_cortInfo+str(iditem)).fetchall()
        return data 

    def obterCliente(self, pedido):
        cliente = self.cursor.execute(self.sql_cliente+pedido).fetchval()
        return cliente

    #ACESSÓRIOS
    def obterIdProd(self, iditem):
        idprod = self.cursor.execute(self.sql_idprod+str(iditem)).fetchval()
        return idprod 

    def obterQuant(self, iditem):
        quant = self.cursor.execute(self.sql_quant+str(iditem)).fetchval()
        return quant
    
    def obterDesc(self, idprod):
        desc = self.cursor.execute(self.sql_desc+str(idprod)).fetchval()
        return desc

    def obterIdGuia(self, pedido):
        idGuia = self.cursor.execute(self.sql_IdGuia+pedido).fetchall()
        return idGuia
    
    def obterTxtGuia(self, idGuia):
        txtGuia = self.cursor.execute(self.sql_txtGuia+str(idGuia)).fetchval()
        return txtGuia

    def obterUnid(self, idprod):
        unid = self.cursor.execute(self.sql_unid+str(idprod)).fetchval()
        return unid

    #PRODUTOS
    def obterPedido(self, iditem):
        pedido = self.cursor.execute(self.sql_pedido+str(iditem)).fetchval()
        return pedido

    def obterListaIDs(self, pedido):
        #Obtendo a lista de id's
        iditens = self.cursor.execute(self.sql_iditem+pedido).fetchall()
        return iditens
   
    def obterLargura(self, iditem):
        largura = self.cursor.execute(self.sql_largura+" t_os_desc_med.IDitem = "+str(iditem)).fetchval()
        return largura

    def obterAltura(self, iditem):
        altura = self.cursor.execute(self.sql_altura+" t_os_desc_med.IDitem = "+str(iditem)).fetchval()
        return altura
    
    def obterItem(self, iditem):
        item = self.cursor.execute(self.sql_item+" t_os_desc.IDitem = "+str(iditem)).fetchval()
        return item    

    def obterValorPedido(self, pedido):
        valor = self.cursor.execute(self.sql_vlrPedido+str(pedido)).fetchval()
        return valor

    def obterObs(self, iditem):
        obs = self.cursor.execute(self.sql_obs+str(iditem)).fetchval()
        return obs

    def obterObsPedido(self, pedido):
        obs = self.cursor.execute(self.sql_obsPedido+str(pedido)).fetchval()
        obs = obs.split("\n")
        email = ""
        for i in range(len(obs)):
            if "@" in obs[i]:
                email = obs[i]

        email = str(const.MAIN_EMAIL) if email == "" else email
        return email

    def obterInfoCliente(self, pedido):
        idcad = self.cursor.execute(self.sql_idCliente+str(pedido)).fetchval()
        all = self.cursor.execute(self.sql_infoCliente+str(idcad)).fetchall()
        cnpj = all[0][0]
        cep = all[0][1]
        end = all[0][2]
        cid = all[0][3]
        bai = all[0][4]
        uf  = all[0][5]
        info = {'cnpj':f'{cnpj}', 'CEP':f'{cep}', 'end':f'{end}', 'localidade':f'{cid}', 'bairro':f'{bai}', 'uf':f'{uf}'}
        return info

class ConnDbConfig:
    def __init__(self):
        #Conexão com a tabela do Sysop
        self.conn = pyodbc.connect(r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
                                   f"DBQ={const.DB_CONFIG_PATH};")
        self.cursor = self.conn.cursor()

        self.sql_peso       = "SELECT peso FROM pesos WHERE pesos.desc LIKE "
        self.sql_pesoById   = "SELECT peso FROM pesos WHERE pesos.id = "
        self.sql_listaOp    = "SELECT desc FROM pesos WHERE pesos.tipo LIKE "
        self.sql_updMsg     = "UPDATE config SET config.desc = "
        self.sql_slcMsg     = "SELECT desc FROM config WHERE config.id = 1"
        self.sql_insPeso    = "INSERT INTO pesos (tipo, desc, peso) VALUES ("
        self.sql_listaIt    = "SELECT desc FROM pesos"
        self.sql_getCond    = "SELECT vlr FROM config WHERE config.id = "
        self.sql_updMaxTres = "UPDATE config SET config.vlr = "
        self.sql_updDoisJun = "UPDATE config SET config.vlr = "

    def obterPeso(self, desc):
        peso = self.cursor.execute(self.sql_peso+"'"+desc+"'").fetchval()
        return peso
    
    def obterPesobyId(self, id):
        peso = self.cursor.execute(self.sql_pesoById+str(id)).fetchval()
        return peso
    
    def obterOpcoes(self, tipo):
        lista = self.cursor.execute(self.sql_listaOp+"'"+tipo+"'").fetchall()
        return lista
    
    def updateObs(self, msg):
        self.cursor.execute(self.sql_updMsg+"'"+msg+"'"+"WHERE config.id = 1")
        self.cursor.commit()
    
    def insertPeso(self, tipo, desc, peso):
        id = self.cursor.execute("SELECT COUNT(*) FROM pesos").fetchval() + 4
        print(str(id)+" "+desc+" "+peso+" "+tipo)
        self.cursor.execute("INSERT INTO pesos VALUES ('"+str(id)+"', '"+desc+"', '"+peso+"', '"+tipo+"')")
        self.cursor.commit()

    def updatePeso(self, desc, peso):
        peso = peso.replace(",",".")
        self.cursor.execute("UPDATE pesos SET pesos.peso = "+peso+" WHERE pesos.desc LIKE '"+desc+"'")
        self.cursor.commit()

    def obterObs(self):
        msg = self.cursor.execute(self.sql_slcMsg).fetchval()
        return msg
    
    def getCondMaxTres(self):
        cond = self.cursor.execute(self.sql_getCond+"3").fetchval()
        return cond

    def getCondDoisJun(self):
        cond = self.cursor.execute(self.sql_getCond+"4").fetchval()
        return cond

    def obterListaItens(self):
        lista = self.cursor.execute(self.sql_listaIt).fetchall()
        return lista
