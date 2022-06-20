import constants as const
from classes.produto import *
from classes.connDatabase import *
from modules.calculation import *
import pandas as pd
import xlsxwriter

#Classe voltada para o relatório

class Relatorio():
    def __init__(self):
        #Iniciando conexão com o banco de dados
        self.conexao = ConnDatabase()
        self.connDbConf = ConnDbConfig()

        #Criando produtos
        self.pr = AllProdutos() #Instância da classe que guarda os produtos

        self.dfProdutos = []

        self.breakrow   = 0
        
        # Obtendo configurações do arquivo json
        with open("C:\SysOp\Reflexa\ROB\expedicao\src\config.json") as f:
            self.config = json.load(f)

    def criarRelatorio(self, pedido, tipo):

        self.dfProdutos = self.selectInfo(pedido)
        
        df = self.dfProdutos

        if tipo == "Embalagem":
            self.insertInfo(pedido, df)
            self.adicionarEstilo(pedido, len(df), len(list(df['Dimensao'].unique()))+3)

        elif tipo == "Peso D":
            self.dfToExcel(df)

        elif tipo == "Peso R":
            df = calcPeso(df)
            df = df.drop(columns=['IdItem', 'Volume', 'Dimensao', 'Colecao', 'Cor', 'Acionamento', 'Tubo', 'Perfil', 'Quantidade', "Tipo"])
            indexDelRows = list(df[df['Nome'] != 'Rolo'].index)
            df = df.drop(indexDelRows) 
            self.dfToExcel(df)

        return df
        

    def editarRelatorio(self, nItem, newVol, pedido):
        #OBJETIVO: dado o número de um item e seu novo volume, o método fará a alteração da informação além de ordenar
        #os itens a partir do volume.

        df = []
        changed = False
        itens = nItem.split(";")

        #Alterar a informação no item
        for l in range(len(itens)):
            i = 0
            found = False
            while not found and i < self.pr.numProdutos:
                if int(itens[l]) == self.pr.lista_Produtos[i].item:
                    self.pr.lista_Produtos[i].volume = newVol
                    self.pr.lista_Produtos[i].pos = 0
                    aux = self.pr.lista_Produtos[i]
                    self.pr.lista_Produtos.pop(i)
                    found = True
                    vol1 = self.pr.lista_Produtos[i].volume
                    vol2 = self.pr.lista_Produtos[i - 1].volume

                    #Verificando se não é o último volume
                    if vol1 != vol2 and vol1 != vol2 + 1 and i != 0:
                        j = i
                        for k in range(j, self.pr.numProdutos - 1):
                            self.pr.lista_Produtos[k].volume -= 1
                        self.pr.numVolumes -= 1
                i += 1    

            #Ordenar os itens a partir do volume
            i = 0
            changed = False
            found = False
            while not changed:
                vol = self.pr.lista_Produtos[i].volume
                if vol == newVol:
                    found = True

                if found:
                    #Caso seja o último elemento da lista
                    if self.pr.numProdutos == i:
                        self.pr.lista_Produtos.insert(i+1, aux)
                        changed = True
                    else:
                        if self.pr.numProdutos > i + 1:
                            if self.pr.lista_Produtos[i - 1].volume != newVol:
                                self.pr.lista_Produtos.insert(i, aux)

                                #Ordenando o item dentro do novo volume
                                j = i
                                while j < self.pr.numProdutos and self.pr.lista_Produtos[j].volume == newVol:
                                    k = j

                                    while k < self.pr.numProdutos and self.pr.lista_Produtos[k + 1].volume == newVol:
                                        if self.pr.lista_Produtos[k].item > self.pr.lista_Produtos[k + 1].item and self.pr.lista_Produtos[k].item == self.pr.lista_Produtos[k + 1].item:
                                            self.pr.lista_Produtos[k], self.pr.lista_Produtos[k + 1] = self.pr.lista_Produtos[k + 1], self.pr.lista_Produtos[k]
                                        k += 1
                                    j += 1
                                changed = True
                i += 1

        self.selectCaixas(df)

        self.pr.lisVolumes = []
        for i in range(1, self.pr.numVolumes + 1):
            dimVol = self.dimVolume(i)
            self.pr.lisVolumes.append(dimVol)

        self.pr.lisDifVol =  list(dict.fromkeys(self.pr.lisVolumes))

        
        self.insertInfo(pedido, df)

        self.adicionarEstilo(pedido, len(df), len(list(df['Dimensao'].unique()))+3)

        itens = []
        for i in range(len(self.pr.lista_Produtos)):
            item =[self.pr.lista_Produtos[i].idItem, self.pr.lista_Produtos[i].desc, self.pr.lista_Produtos[i].volume, self.pr.lista_Produtos[i].dim, self.pr.lista_Produtos[i].item, self.pr.lista_Produtos[i].largura, self.pr.lista_Produtos[i].altura, self.pr.lista_Produtos[i].modelo, self.pr.lista_Produtos[i].colecao, self.pr.lista_Produtos[i].cor, self.pr.lista_Produtos[i].acion, self.pr.lista_Produtos[i].quant, self.pr.lista_Produtos[i].tubo, self.pr.lista_Produtos[i].perfil, self.pr.lista_Produtos[i].altCom, self.pr.lista_Produtos[i].tipo, self.pr.lista_Produtos[i].pos, 0]
            itens.append(item)
        
        # Criando dataframe 
        df = pd.DataFrame(data=itens, columns=["IdItem", "Nome", "Volume", "Dimensao", "Item", "Largura", "Altura", "Modelo", "Colecao", "Cor", "Acionamento", "Quantidade", "Tubo", "Perfil", "Altura do Comando", "Tipo", "Posicao", "Peso"])
        return df


    def selectInfo(self, pedido):
        #Método fará a busca por todas as informações necessárias
        #Obtendo o número de produtos de um pedido
        
        df = self.getItens(pedido)
        df = self.selectVolumes(df)
        df = self.selectCaixas(df)

        self.pr.lisVolumes = []
        for i in range(1, self.pr.numVolumes + 1):
            dimVol = self.dimVolume(i)
            self.pr.lisVolumes.append(dimVol)

        self.pr.lisDifVol =  list(dict.fromkeys(self.pr.lisVolumes))

        return df


    def getItens(self, pedido):
        #OBJETIVO: Fará a busca e atribuição das informações de todos os itens de um pedido
        # pedido: String

        produtos = 0
        acessorios = 0
        listaIds = []
        self.pr.numProdutos = 0
        pedido = pedido.split(';')

        for i in range(len(pedido)):        
            #Obter lista de id's
            idItens = self.conexao.obterListaIDs(pedido[i])
            self.pr.numProdutos += len(idItens)

            for j in range(len(idItens)):
                listaIds.append(idItens[j][0])

        #Buscando e excluíndo, caso exista, o registro sobre o serviço de instalação
        i = 0
        while (i < self.pr.numProdutos):
            iditem = listaIds[i]
            idprod  =   self.conexao.obterIdProd(iditem)

            if idprod < 15:
                produtos += 1
            else: acessorios += 1

            if idprod == 1462 or idprod == 2230:
                listaIds.remove(iditem)
                self.pr.numProdutos = self.pr.numProdutos - 1
                i -= 1
            i += 1
        
        #Buscando todas as informações dos produtos
        hasMulti = False
        x = 0
        self.pr.lista_Produtos = []
        for iditem in listaIds:
            idprod =    self.conexao.obterIdProd(iditem)
            tipo    =   "Produto" if idprod < 15 and idprod != 0 else "Acessorio"
            desc =      self.conexao.obterDesc(idprod)

            if "SERVIÇO" not in desc.upper():
                self.pr.lista_Produtos.append(Produto())

                if idprod > 15 and x < produtos: #Caso o produto seja um acessório e não esteja no final da lista, mandá-lo para o fim da lista
                    listaIds.append(listaIds.pop(listaIds.index(iditem)))
                    x -= 1
                else:   
                    self.pr.lista_Produtos[x].changeIdItem(iditem)

                if tipo == "Produto":
                    self.pr.lista_Produtos[x] = self.createProduto(iditem, idprod, tipo)
                    if containsWord(self.pr.lista_Produtos[x].modelo, "Mult"):
                        hasMulti = True
                        largs = self.conexao.obterObs(iditem).split("|")
                        for j in range(len(largs)):
                            if j == 0:
                                self.pr.lista_Produtos[x].largura = float(largs[j].replace(",", "."))
                            else:
                                self.pr.lista_Produtos.append(self.createProduto(iditem, idprod, tipo, largs[j])) 
                                produtos += 1
                                x += 1
                                self.pr.numProdutos += 1

                elif tipo == "Acessorio":
                    quant   = self.conexao.obterQuant(iditem)
                    ped     = self.conexao.obterPedido(iditem)
                    desc    = self.conexao.obterDesc(idprod)
                    #ATRIBUIÇÕES-------------------------------------------------------------------------------------------
                    self.pr.lista_Produtos[x].changeQuant(str(int(quant)))
                    self.pr.lista_Produtos[x].changePedido(ped)
                    self.pr.lista_Produtos[x].changeDesc(desc)
                    self.pr.lista_Produtos[x].changeTipo(tipo)
            else:
                self.pr.numProdutos -= 1
                listaIds.pop(listaIds.index(iditem))
                x -= 1
            x += 1
                

        if hasMulti:
            i = 0
            while i < len(self.pr.lista_Produtos):
                tipo = self.pr.lista_Produtos[i].tipo
                if tipo == "Acessorio" and i < produtos: #Caso o produto seja um acessório e não esteja no final da lista, mandá-lo para o fim da lista
                    self.pr.lista_Produtos.append(self.pr.lista_Produtos.pop(self.pr.lista_Produtos.index(self.pr.lista_Produtos[i])))
                    i -= 1
                i += 1

        #Ordenar lista de produtos pelo item
        changed = True
        while changed:
            changed = False
            k = 0
            while k + 1 < produtos and self.pr.lista_Produtos[k + 1].tipo == "Produto":
                if self.pr.lista_Produtos[k].item > self.pr.lista_Produtos[k + 1].item:
                    self.pr.lista_Produtos[k], self.pr.lista_Produtos[k + 1] = self.pr.lista_Produtos[k + 1], self.pr.lista_Produtos[k]
                    changed = True
                k += 1
            j += 1
        
        #Verificando se há acessórios iguais
        i = produtos
        j = 0
        while i < len(self.pr.lista_Produtos):
            aux = False
            j = i
            nome = self.pr.lista_Produtos[i].desc

            while j + 1 < len(self.pr.lista_Produtos) and not aux: 
                if nome == self.pr.lista_Produtos[j + 1].desc:
                    quant01 = int(self.pr.lista_Produtos[i].quant) if self.pr.lista_Produtos[i].quant != "" else 0
                    quant02 = int(self.pr.lista_Produtos[j + 1].quant) if self.pr.lista_Produtos[j + 1].quant != "" else 0
                    self.pr.lista_Produtos[i].quant = quant01 + quant02
                    self.pr.lista_Produtos.remove(self.pr.lista_Produtos[j + 1])
                    self.pr.numProdutos -= 1
                    aux = True
                j += 1
            i += 1
        #Adicionando o conjunto de instalação na lista de acessórios
        lisConj = self.createConjIns()
        for i in range(len(lisConj)):
            self.pr.lista_Produtos.append(lisConj[i])

            #Verifica se o conjunto é da QS85, caso seja, adicionar o item 'clipe'
            if containsWord(lisConj[i].desc, "ACMEDA/FASCIA"):
                newP = Produto()
                newP.quant = lisConj[i].quant * 2
                newP.desc = "CLIPE FASCIA BRANCA"
                newP.tipo = "Acessorio"
                self.pr.lista_Produtos.append(newP)
                self.pr.numProdutos += 1
            self.pr.numProdutos += 1

        #Obter informações sobre guia, tubo afastador...
        self.createGuias(pedido[0])

        # #Obter cliente
        self.pr.cliente = self.conexao.obterCliente(pedido[0]) 
        if(self.pr.cliente is None): #Caso não haja a informação do cliente, tornar o atributo em um espaço vazio
            self.pr.cliente = " "

        itens = []
        for i in range(len(self.pr.lista_Produtos)):
            item =[self.pr.lista_Produtos[i].idItem, self.pr.lista_Produtos[i].desc, self.pr.lista_Produtos[i].volume, self.pr.lista_Produtos[i].dim, self.pr.lista_Produtos[i].item, self.pr.lista_Produtos[i].largura, self.pr.lista_Produtos[i].altura, self.pr.lista_Produtos[i].ambiente, self.pr.lista_Produtos[i].modelo, self.pr.lista_Produtos[i].colecao, self.pr.lista_Produtos[i].cor, self.pr.lista_Produtos[i].acion, self.pr.lista_Produtos[i].quant, self.pr.lista_Produtos[i].tubo, self.pr.lista_Produtos[i].perfil, self.pr.lista_Produtos[i].altCom, self.pr.lista_Produtos[i].tipo, self.pr.lista_Produtos[i].pos, 0]
            itens.append(item)
        
        # Criando dataframe 
        df = pd.DataFrame(data=itens, columns=["IdItem", "Nome", "Volume", "Dimensao", "Item", "Largura", "Altura", "Ambiente", "Modelo", "Colecao", "Cor", "Acionamento", "Quantidade", "Tubo", "Perfil", "Altura do Comando", "Tipo", "Posicao", "Peso"])
        return df


    def createProduto(self, iditem, idprod, tipo, lar = 0):
        #OBJETIVO: Fará a criação de um produto, buscando as informações no banco de dados
        newProd = Produto()

        infoCortina = self.conexao.obterCortInfo(iditem)
        # info Cortina -> resultado de query que obtém informações da cortina
        # 0: Cliente
        # 1: Ambiente
        # 2: Modelo
        # 3: Coleção
        # 4: Cor
        # 5: Cor Fabricante
        # 6: Acionamento
        # 7: Tubo
        # 8: Observações
        # 9: Perfil
        # 10: Altura do Comando

        item =      self.conexao.obterItem(iditem)
        pedido =    self.conexao.obterPedido(iditem)
        largura =   self.conexao.obterLargura(iditem) if lar == 0 else float(lar.replace(",", "."))
        altura =    self.conexao.obterAltura(iditem)
        desc =      self.conexao.obterDesc(idprod)
        tubo =      infoCortina[0][7] if infoCortina[0][7] is not None else ""
        perfil =    infoCortina[0][9] if infoCortina[0][9] is not None else ""

        newProd.idItem = iditem
        newProd.idProd = idprod
        newProd.changeItem(item) 
        newProd.changeIdProd(idprod) 
        newProd.changePedido(pedido)
        newProd.changeTipo(tipo)
        newProd.changeAmbiente(infoCortina[0][1])
        newProd.changeLargura(largura)
        newProd.changeAltura(altura)
        newProd.changeModelo(infoCortina[0][2])
        newProd.changeColecao(infoCortina[0][3])
        newProd.changeCor(infoCortina[0][5])
        newProd.changeAcion(infoCortina[0][6])
        newProd.changeDesc(desc)
        newProd.changeTubo(tubo)
        newProd.changeAltCom(infoCortina[0][10])
        newProd.changePerfil("Perfil "+ perfil)

        return newProd


    def dfToExcel(self, df):
        #OBJETIVO: Dado um dataframe, ele fará a inserção das informações em um arquivo .xlsx
        # df : Dataframe

        path = selectPath("relatorio_peso")

        writer = pd.ExcelWriter(path, engine='xlsxwriter')

        df.to_excel(writer, sheet_name='Sheet1', startrow=1, header=False, index=False)

        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']

        header_format = workbook.add_format({
            'bold': True,
            'fg_color': '#9BA4B4',
            'font_color': 'white',
            'border': 1})

        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)

            column_len = df[value].astype(str).str.len().max()
            column_len = max(column_len, len(value))

            worksheet.set_column(col_num, col_num, column_len)

        bg_format1 = workbook.add_format({'bg_color': '#EEEEEE'}) # white cell background color
        bg_format2 = workbook.add_format({'bg_color': '#FFFFFF'}) # white cell background color

        for i in range(len(df)): # integer odd-even alternation 
            worksheet.set_row(i, cell_format=(bg_format1 if i%2!=0 else bg_format2))

        writer.save()


    def insertInfo(self, pedido, df):
        #OBJETIVO: Método fará a inserção de todos os dados no relatório
        #pedido: String

        difVol = list(df['Volume'].unique())
        difDim = list(df['Dimensao'].unique())
        dims = []
        Ppartida = len(difDim)+3   #Ponto de partida para inserir as informações será relativo à quantidade de dimensões
        Ppartida = Ppartida if Ppartida >= 6 else 6
        path = selectPath(pedido)
        df = calcPeso(df)
        self.pesos = {}
         

        # Obtendo o peso por volume
        self.pesos = {i: (round(float(df.loc[df['Volume'] == i, ['Peso']].sum()) * 100) / 100) for i in difVol}
        df.reset_index(inplace=True)
        indexDelRows = list(df[df['Tipo'] == 'Caixa'].index)
        df = df.drop(indexDelRows)

        pesoTotal = {0:round(sum(self.pesos.values())*100)/100}
        pesoTotal.update(self.pesos)
        
        #Leitura do modelo
        workbook = xlsxwriter.Workbook(path)
        worksheet = workbook.add_worksheet()

        #Informações imutáveis
        worksheet.write("A1", "PEDIDO")
        worksheet.write('B1', 'VALORES')
        worksheet.write('C1', 'QUANTIDADE')
        worksheet.write('D1', 'TAMANHO DA CAIXA')

        worksheet.write('A2', pedido + " - " + self.pr.cliente)
        worksheet.write('A3', 'VOLUME')
        worksheet.write('A4', 'PESO(kg)')

        #Adicionando o peso do pedido
        worksheet.write('B4', pesoTotal[0])

        #Adicionar a quantidade de volumes no pedido
        worksheet.write('B3', len(difVol))

        for vol in difVol:
            dims.append((df.loc[df['Volume'] == vol, 'Dimensao']).iloc[0])

        #Adicionando quantidade de volumes de cada dimensão
        for i in range(1, len(difDim)+1):   
            worksheet.write(i, 3, difDim[i-1])  
            worksheet.write(i, 2, dims.count(difDim[i - 1]))

        #PRODUTOS - Adicionando informações gerais
        x = 0
        for id in df.index:
            row = Ppartida + x

            # Verificar breakrow da página:
            # Dessa forma, a página não será impressa com parte de um volume em uma página
            # e outra parte em outra página.
            # Com o ajuste de página em 93%, a linha limite é 55
            if row == 55 and id != df.index[-1]:
                ind = list(df.index).index(id)
                notDif = True

                while notDif:
                    vol1 = df["Volume"][df.index[ind]]
                    vol2 = df["Volume"][df.index[ind + 1]]

                    if vol1 != vol2:
                        self.breakrow = row + 1
                        notDif = False
                    row -= 1
                    ind -= 1

            #Informações da coluna 'pedido'
            worksheet.write(Ppartida + x, 0, df["Volume"][id])

            #Informações da coluna 'valores'
            v = int(df["Volume"][id])
            pos = df["Posicao"][id]

            if self.mustShow(pos, v):
                worksheet.write(Ppartida + x, 1, df["Dimensao"][id] + " | " + str(pesoTotal[v]) + "Kg")
            else:
                worksheet.write(Ppartida + x, 1, "")

            #Informações da coluna 'quantidade'
            if df["Tipo"][id] == "Produto":
                quantidade = df["Item"][id]
            else:
                quantidade = df["Quantidade"][id]

            worksheet.write(Ppartida + x, 2, quantidade)

            #Informações da coluna 'tamanho da caixa'
            if df['Tipo'][id] == "Produto":
                lar = "{:.3f}".format(df["Largura"][id])
                alt = "{:.3f}".format(df["Altura"][id])

                lar = lar.replace(".", ",")
                alt = alt.replace(".", ",")
                output = df["Nome"][id]+"("+df["Ambiente"][id]+") "+lar+" X "+alt
            else:
                output = df["Nome"][id]

            worksheet.write(Ppartida + x, 3, output)        

            x += 1    

        workbook.close()


    def adicionarEstilo(self, pedido, np, Ppartida):
        #OBJETIVO: Adicionar estilo ao arquivo xlsx.
        # Pedido
        # np : Integer -> Número de produtos
        # Ppartida : Integer -> Valor de início de alterações no documento 

        Ppartida = Ppartida if Ppartida >= 6 else 6
        path = selectPath(pedido)

        #ESTILO DA TABELA ----------------------------------------------------------------------------------------
        table1 = pd.read_excel(path)

        writer = pd.ExcelWriter(path, engine='xlsxwriter')
        table1.to_excel(writer, index=False, sheet_name='report')

        workbook = writer.book
        worksheet = writer.sheets['report']
        
        
        worksheet.set_margins(left=0.3, right=0.3, top=0.3, bottom=0.3)
        worksheet.set_print_scale(93)
        if self.breakrow != 0:
            worksheet.set_h_pagebreaks([self.breakrow])
            self.breakrow = 0

        
        #Declarando formatações
        volume_fmt = workbook.add_format({'align': 'center'}) 
        itens_fmt = workbook.add_format({'align': 'center'})
        pedido_fmt = workbook.add_format({'align': 'center', 'border': 1})
        oddRow  = workbook.add_format({'bg_color': '#808080'})

        titulo_row = workbook.add_format(
            {'align': 'center',
            'font_name': 'Arial',
            'font_color': 'white',
            'bold': 'True',
            'bg_color': 'black'})

        #Alterando o zoom do relatório
        worksheet.set_zoom(75)

        #Delimitar área
        fimRows = str(np+9)
        worksheet.conditional_format( 'C'+str(Ppartida+1)+':D'+fimRows , { 'type' : 'no_blanks' , 'format' : pedido_fmt} )
        worksheet.conditional_format( 'A'+str(Ppartida+1)+':A'+fimRows , { 'type' : 'no_blanks' , 'format' : pedido_fmt} )
        
        #Alterando o tamanho das colunas
        worksheet.set_column('A:A', 20, volume_fmt)
        worksheet.set_column('B:B', 20, itens_fmt)
        worksheet.set_column('C:C', 15, itens_fmt)
        worksheet.set_column('D:D', 50)

        #Primeira linha de títulos
        worksheet.write("A1", "PEDIDO", titulo_row)
        worksheet.write("B1", " ", titulo_row)
        worksheet.write("C1", "QUANTIDADE", titulo_row)
        worksheet.write("D1", "TAMANHO DA CAIXA", titulo_row)

        #Segunda linha de títulos
        worksheet.write("A"+str(Ppartida), "VOLUME", titulo_row)
        worksheet.write("B"+str(Ppartida), "TAMANHO DA CX", titulo_row)
        worksheet.write("C"+str(Ppartida), "ITEM", titulo_row)
        worksheet.write("D"+str(Ppartida), pedido+" - "+self.pr.cliente, titulo_row)

        #Alterar cor de fundo baseado no volume
        delimitador = 'A'+str(Ppartida+1)+':D'+str(np+Ppartida)
        worksheet.conditional_format(delimitador, 
                                    {'type':     'formula',
                                     'criteria': '=MOD($A'+str(Ppartida+1)+', 2)=0',
                                    'format':   oddRow})

        writer.save()


    def mustShow(self, pos, v):
        #OBJETIVO: dirá se algumas informações devem ser postas no relatório. Configuração puramente
        # visual.

        quant   = 0 #Quantidade de produtos no volume
        existe  = False

        #Buscar a quantidade de produtos no volume
        for i in range(len(self.pr.lista_Produtos)):
            if self.pr.lista_Produtos[i].volume == v:
                quant = quant + 1
        
        #Pesquisando se há a posição quant/2
        for i in range(len(self.pr.lista_Produtos)):
            if self.pr.lista_Produtos[i].pos == (quant/2) and self.pr.lista_Produtos[i].volume == v:
                existe = True
                
        if pos == (quant/2) or (not existe and pos == (quant/2) + 0.5 ):
            return True
        else: return False
        

    def createGuias(self, pedido):
        #OBJETIVO: Fará a criação das guias e adicionará na lista de produtos.

        ids = self.conexao.obterIdGuia(pedido)          #Lista dos ids dos itens e o id da guia de cada item
        guias = [Produto() for i in range(len(ids))]    #Lista de produtos que conterá, temporariamente as guias
        descricoes = []                                 #Lista das descrições de cada acessório encontrado

        #Criação dos acessórios
        for i in range(len(ids)):
            desc = self.conexao.obterTxtGuia(ids[i][1]).upper()
            

                
            if "GUIA" in desc:
                if containsWord(desc, "2X"):
                    desc = desc.replace("2X", "")
                elif containsWord(desc, "BR"):
                    desc = desc.replace("BR", "BRANCA")
                alt = self.conexao.obterAltura(ids[i][0])
                guias[i].altura = alt
                guias[i].desc = desc+" - "+str(int(alt*1000))+"MM"
                guias[i].quant = 2
            else:
                alt = self.conexao.obterAltura(ids[i][0])
                guias[i].altura = alt
                guias[i].desc = desc
                guias[i].quant = 1

            descricoes.append(guias[i].desc)

        #Obter somente os acessórios diferentes
        descricoes = list(dict.fromkeys(descricoes))

        #Criará os produtos que será de fato inseridos na lista principal
        for i in range(len(descricoes)):
            newp = Produto()
            newp.quant = 0
            newp.desc = descricoes[i]
            newp.tipo = "Acessorio"
            
            #Obter a quantidade 
            for j in range(len(ids)):
                if guias[j].desc == newp.desc:
                    newp.quant = newp.quant + guias[j].quant
                    newp.altura = guias[j].altura

            self.pr.lista_Produtos.append(newp)
            self.pr.numProdutos = self.pr.numProdutos + 1
            
            #Um volume não pode ter mais de oito guias juntas, logo, se houver mais de oito guias em um mesmo produto
            #deve-se criar novos produtos com o restante das guias
            while self.pr.lista_Produtos[len(self.pr.lista_Produtos)-1].quant > 8:
                    pos = len(self.pr.lista_Produtos)-1         #Posição no vetor
                    quant = self.pr.lista_Produtos[pos].quant   #Quantidade do produto
                    desc = self.pr.lista_Produtos[pos].desc
                    alt  = self.pr.lista_Produtos[pos].altura

                    self.pr.lista_Produtos[pos].quant = 8 

                    newGuia = Produto()
                    newGuia.quant = quant - 8
                    newGuia.desc  = desc
                    newGuia.tipo  = "Acessorio"
                    newGuia.altura= alt

                    self.pr.lista_Produtos.append(newGuia)
                    self.pr.numProdutos = self.pr.numProdutos + 1


    def createConjIns(self):
        #OBJETIVO: Fará a seleção do conjunto de instalação de um pedido

        conj = [Produto()]    #Conjunto de instalação
        comp = ""           #Componente
        quant = 0           #Quantidade
        np = 0
       
        #Verificando a quantidade de produtos (retirando os acessórios)
        for l in range(len(self.pr.lista_Produtos)):
            if self.pr.lista_Produtos[l].tipo == "Produto":
                np = np + 1

        #Fará a contagem de produtos
        for i in range(len(self.pr.lista_Produtos)):
            if self.pr.lista_Produtos[i].tipo == "Produto":
                modelo = self.pr.lista_Produtos[i].modelo
                item = self.pr.lista_Produtos[i].item
                cor = self.pr.lista_Produtos[i].cor
                cor = cor.upper() if cor is not None else ""
                supInt = 0            

                #Selecionando os componentes
                comp = selectComp(modelo) + " " + cor

                #Calculando a quantidade do componente
                if containsWord(modelo, "Mult") and i != 0:             #Caso seja mult link, devemos contabilizar a cada número de item
                    if item != self.pr.lista_Produtos[i - 1].item:
                        quant = calcQuantConj(self.pr.lista_Produtos[i].largura, comp)
                        supInt = calcSupInt(self.pr.lista_Produtos[i].modelo)
                    else:
                        quant = 0
                else:
                    quant = calcQuantConj(self.pr.lista_Produtos[i].largura, comp)


                for j in range(len(conj)):
                    nomes = [conj[x].desc for x in range(len(conj))]
                    if conj[j].desc == comp:
                        conj[j].quant += quant
                        conj[j].supInt += supInt
                        conj[j].tipo = "Acessorio"
                        conj[j].modelo = modelo

                    elif conj[j].desc == "":
                        conj[j].desc = comp 
                        conj[j].quant = quant
                        conj[j].supInt = calcSupInt(self.pr.lista_Produtos[i].modelo)
                        conj[j].tipo = "Acessorio"
                        conj[j].modelo = modelo

                    elif comp not in nomes: 
                        newC = Produto()
                        newC.desc   = comp
                        newC.quant  = quant
                        newC.supInt = calcSupInt(self.pr.lista_Produtos[i].modelo)
                        newC.tipo   = "Acessorio"
                        conj[j].modelo = modelo
                        conj.append(newC)

        for i in range(len(conj)):
            if conj[i].supInt != 0:
                newC = Produto()
                newC.desc = "SUP. INTERMEDIÁRIO"
                newC.quant = conj[i].supInt
                newC.tipo = "Acessorio"
                newC.modelo = conj[i].modelo
                conj.append(newC)
            
        return conj


    def dimVolume(self, v):
        #OBJETIVO: dado um volume, o método retorna a sua dimensão.

        a = ""  #Variável auxiliar
        i = 0   #Variável de contagem

        while a == "":
            vol = self.pr.lista_Produtos[i].volume
            if vol == v:
                a = self.pr.lista_Produtos[i].dim
            i = i + 1
        return a


    def selectVolumes(self, df):
        #Método determinará o volume de cada item

        np = len(df)    #Número de produtos
        i = 0                       #Contagem de elementos de um volume
        v = 1                       #Volume
        cheio   = False             #Indica se o volume está cheio
        larS    = 0                 #Resultado do método larSuficiente
        condDoisJun = self.config['doisJuntos']
        z = 0
        for id in df.index:
            cheio = self.isVolFull(i,z)
            larS = larSuficiente(df['Largura'][id], condDoisJun)

            if cheio: 
                v = v + 1
                i = 0 
            
            # Adicionando informação do volume no objeto e no dataframe
            self.pr.lista_Produtos[z].volume = v
            df.loc[id, ['Volume']] = v

            i = i + larS
            df.loc[id, ['Posicao']] = i
            self.pr.lista_Produtos[z].pos = i
            z += 1
            
        self.quantVolumes()

        return df


    def isVolFull(self, p, i):
        #OBJETIVO: Dada a posição do produto no volume e o 
        #volume, o método indicará se o produto pode pertencer ao volume
        #p: float   - Posição do produto no volume
        #v: Integer - Volume

        pesado = False  #Caso algum dos produtos tenha a largura maior que 1.85m o volume se torna pesado
        condDoisJun = self.config['doisJuntos']
        l      = larSuficiente(self.pr.lista_Produtos[i].largura, condDoisJun)
        isGuia  =  True if "GUIA" in self.pr.lista_Produtos[i].desc else False
        cond    = self.config['maxTres']
        a      = 0      #Auxiliar
        b      = 0      #Auxiliar

        #Caso seja um acessório
        if self.pr.lista_Produtos[i].tipo != "Acessorio":
            if p <= 2 or (p == 2.5 and l != 1):
                pesado = False
            else: 
                #Verificar se o produto e todos os outras antes dele são menores que 1.85
                a = p + l
                while a != 0:
                    if self.pr.lista_Produtos[i - b].largura >= cond:
                        pesado = True
                    a = a - larSuficiente(self.pr.lista_Produtos[i - b].largura, condDoisJun)
                    b = b + 1
                
                #Verificar se todos os produtos que virão após o produto serão leves
                a = p + l
                b = 1
                l = 0
                while (a < 6) and (False if a == 5.5 and l == 1 else True) and (i + b) <= len(self.pr.lista_Produtos) - 1:
                    if self.pr.lista_Produtos[i + b].largura >= cond:
                        pesado = True

                    l = larSuficiente(self.pr.lista_Produtos[i + b].largura, condDoisJun)
                    b = b + 1
                    a = a + l

        else:
            if self.pr.lista_Produtos[i - 1].tipo == "Acessorio":
                pesado = False
            else:
                quant = 0
                for j in range(len(self.pr.lista_Produtos)):
                    if self.pr.lista_Produtos[j].tipo == "Produto":
                        quant += 1
                if quant >= 40:
                    pesado = True

        if pesado and (True if p + larSuficiente(self.pr.lista_Produtos[i].largura, condDoisJun) > 3 else False): return True
        elif not pesado and (True if p + larSuficiente(self.pr.lista_Produtos[i].largura, condDoisJun) > 6 else False): return True
        elif isGuia: return True
        else: return pesado


    def selectCaixas(self, df):
        #OBJETIVO: Fará a seleção da caixa que será utilizada em cada volume

        np = self.pr.numProdutos    #Número de produtos
        nv = self.pr.numVolumes     #Número de volumes
        count = 0                   #Contagem da quantidade de produtos em um volume
        mLar = 0                    #Maior largura de um volume
        vol = 0                     #Volume do produto
        dim = ""                    #Dimensão da caixa

        for i in range(1, nv + 1):
            count = self.countProd(i) 
            mLar = self.maiorLar(i)
            j = 0
            for id in (df.index):
                desc = self.pr.lista_Produtos[j].desc
                tipo = self.pr.lista_Produtos[j].tipo
                alt  = self.pr.lista_Produtos[j].altura
                dim = self.selectDim(mLar, alt, count, desc, tipo)
                vol = self.pr.lista_Produtos[j].volume
                
                if vol == i:
                    self.pr.lista_Produtos[j].changeDim(dim)
                    df.loc[id, ['Dimensao']] = dim
                j += 1

        return df

    def selectDim(self, l, a, n, desc, tipo):
        #OBJETIVO: Dada a maior largura e o número produtos de um volume, retorna as dimensões da caixa
        #l: Double      - Maior largura 
        #n: Integer     - Número de elementos
        #desc: String   - Descrição do produto
        #tipo: String   - Tipo do produto


        lar = l + 0.14      #Espaço a mais
        alt = a + 0.14      #Altura com o espaço a mais. Será usado no cálculo da guia
        comp = ""           #Comprimento vai variar de acordo com o número de peças do volume
        
        if n > 3:
            comp = "270x250"
        else: comp = "220x200"

        if lar <= 1.4 and l != 0:
            return comp+"x1400"
        elif lar > 1.4 and lar <= 1.6:
            return comp+"x1600" 
        elif lar > 1.6 and lar <= 1.8:
            return comp+"x1800"
        elif lar > 1.8 and lar <= 2:
            return comp+"x2000"
        elif l == 0:
            if containsWord(desc, "GUIA") or containsWord(desc, "TUBO"):
                return "270x250x"+str(int(round(alt, 1)*1000))
            else: 
                return "240x300x340"
        else:
            return comp+"x"+str(int(round(lar, 1)*1000))


    def maiorLar(self, v):
        #OBJETIVO: Dado um volume, o método retornará a maior largura presente. Deverá ser considerados os casos em 
        #dois produtos são colocados lado a lado
        #v: Integer - Volume

        np = self.pr.numProdutos    #Número de produtos
        vol = 0                     #Volume do produto
        lar = 0                     #Largura do produto
        result = 0                  #Resultado
        larguras = []               #Todas as larguras 
        menLarguras = []            #Menores larguras, caso sejam menores que 0.760
        a = 0                       #Variável auxiliar para obter maior largura de menLarguras
        acess  = False
        prod    = False
        
        #Atribuindo valores à lista larguras
        for i in range(len(self.pr.lista_Produtos)):
            tipo = self.pr.lista_Produtos[i].tipo
            lar = self.pr.lista_Produtos[i].largura
            vol = self.pr.lista_Produtos[i].volume

            if vol == v:
                if tipo == "Produto":
                    prod = True
                else: 
                    acess = True
                larguras.append(lar)

        #Obtendo a maior largura da lista
        result = max(larguras)

        #Atibuindo valores à lista menLarguras
        menLarguras = list(filter(lambda x: x <= 0.760, larguras))

        if len(menLarguras) > 1:
            a = max(menLarguras)
            menLarguras.remove(a)
            a = a + max(menLarguras)
        
        if a > result:
            result = a
        
        if acess and not prod:
            result = 0

        return result


    def countProd(self, n):
        #Fará a contagem de quantos produtos um volume possui
        #Deverá considerar dois produtos pequenos como um só
        #n: Número do volume
        
        np = self.pr.numProdutos    #Número de produtos
        j = 0                       #Contagem dos produtos
        vol = 0                     #Volume do produto
        lar = 0                     #Largura do volume
        condDoisJun = float(self.connDbConf.getCondDoisJun())

        for i in range(len(self.pr.lista_Produtos)):
            if self.pr.lista_Produtos[i].tipo == "Produto":
                lar = self.pr.lista_Produtos[i].largura
            else:
                lar = 0

            vol = self.pr.lista_Produtos[i].volume

            if vol == n:
                j = j + larSuficiente(lar, condDoisJun)
        return j
        

    def quantVolumes(self):
        #OBJETIVO: Obter a quantidade de volumes no pedido

        listaVolumes = [] #Lista de todos os volumes

        for i in range(len(self.pr.lista_Produtos)):
            listaVolumes.append(self.pr.lista_Produtos[i].volume)
            
        self.pr.numVolumes = max(listaVolumes)    


    def getConjDesc(self):
        #OBJETIVO: Retornará a descrição do conjunto de suporte de instalação

        for i in range(self.pr.numProdutos):
            desc = self.pr.lista_Produtos[i].desc
            if (containsWord(desc, "SUP.") or containsWord(desc, "PRESILHA")):
                return desc
    


