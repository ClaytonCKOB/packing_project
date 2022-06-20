#Classe com todos os produtos de um pedido

class AllProdutos:
    def __init__(self):
        self.lista_Produtos = []
        self.lisVolumes = []    #Lista das dimensões de todos os volumes
        self.lisDifVol  = []    #Lista das dimensões, porém sem repetição
        self.cliente = ""
        self.numProdutos = 0
        self.numVolumes = 0


class Volume:
    def __init__(self, nr, dimensao, peso):
        self.nr = nr
        self.dim = dimensao
        self.peso = peso


#Classe que representará o produto
class Produto:
    def __init__(self):
        self.idItem = ""
        self.idProd = ""
        self.pedido = ""
        self.volume = 0
        self.item   = ""
        self.ambiente = ""
        self.largura = 0
        self.altura = 0
        self.modelo = ""
        self.colecao = ""
        self.cor    = ""
        self.acion  = ""
        self.desc   = ""
        self.dim    = ""
        self.quant  = ""
        self.tubo   = ""
        self.perfil = ""
        self.altCom = 0
        self.supInt = 0     #Suporte intermediário usado nas mult link
        self.tipo   = ""    #Produto ou Acessorio
        self.pos    = 0     #Posição do produto no volume

    def changeTipo(self, tipo):
        self.tipo = tipo

    def changePedido(self, pedido):
        self.pedido = pedido

    def changeQuant(self, quant):
        self.quant = quant

    def changeIdItem(self, id):
        self.idItem = id

    def changeIdProd(self, idprod):
        self.idProd = idprod

    def changeVolume(self, volume):
        self.volume = volume

    def changeItem(self, item):
        self.item = item
    
    def changeAmbiente(self, ambiente):
        self.ambiente = ambiente
    
    def changeLargura(self, largura):
        self.largura = largura
    
    def changeAltura(self, altura):
        self.altura = altura

    def changeModelo(self, modelo):
        self.modelo = modelo

    def changeColecao(self, colecao):
        self.colecao = colecao
    
    def changeCor(self, cor):
        self.cor = cor

    def changeAcion(self, acion):
        self.acion = acion

    def changeDesc(self, desc):
        self.desc = desc

    def changeDim(self, dimensao):
        self.dim = dimensao

    def changeTubo(self, tubo):
        self.tubo = tubo
    
    def changePerfil(self, perfil):
        self.perfil = perfil
    
    def changeAltCom(self, altCom):
        self.altCom = altCom