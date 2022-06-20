#Arquivo voltado para deselvover a função que realizará os cálculos do peso
import pandas as pd
from classes.connDatabase import *
from modules.geralFun import *
import requests


conn = ConnDbConfig()

def calcPeso(df = 0) -> dict:
    #OBJETIVO: Fará o cálculo do peso total de um pedido

    # Variáveis
    peso = 0
    numP = 0
    msg     = ""
    modMagnaBasic = ["QS80", "QS83"]
    modMagnaPlus = ["QS82", "QS81"]
    modMidiCassete = ["R011", "R012"]


    #Obter o número de produtos
    numP = df["Tipo"].value_counts().Produto

    #Varrer cada produto
    for i in (df.index):
        quant   = df['Quantidade'][i]
        altCom  = df['Altura do Comando'][i]
        alt     = df['Altura'][i]
        lar     = df['Largura'][i]
        desc    = df['Nome'][i]
        tecido  = df['Colecao'][i]
        tubo    = df['Tubo'][i]
        perfil  = df['Perfil'][i]
        acion   = df['Acionamento'][i]
        modelo  = df['Modelo'][i]
        
        area    = alt*lar
        peso = 0
            
        #CÁLCULO DO PESO DOS ACESSÓRIOS
        if df['Tipo'][i] == "Acessorio":
            if containsWord(desc, "GUIA"):
                vlrPeso = conn.obterPesobyId(24)
                guia = vlrPeso * alt * quant
                peso += guia
            
            else:
                #CÁLCULO DOS COMPONENTES
                if containsWord(modelo, modMagnaBasic):
                    vlrPeso = conn.obterPesobyId(1)
                    comp = vlrPeso*numP
                    peso += comp

                elif containsWord(modelo, modMagnaPlus):
                    vlrPeso = conn.obterPesobyId(3)
                    comp = vlrPeso*numP
                    peso += comp

                elif containsWord(modelo, modMidiCassete):
                    vlrPeso = conn.obterPesobyId(2)
                    comp = vlrPeso*numP
                    peso += comp

                else:
                    Pace = (conn.obterPeso(desc) if conn.obterPeso(desc) is not None else 0) 
                    if Pace == 0:
                        if not containsWord(desc, "Bucha") and not containsWord(desc, "Fita") and not containsWord(desc, "Parafuso"):
                            msg += "OUTRO: "+desc+"|"
                    peso += Pace

        #CÁLCULO DO PESO DOS PRODUTOS
        else:
            #Calcular peso do tecido
            
            Ptecido = conn.obterPeso(getTecido(tecido))
            Ptecido = (Ptecido if Ptecido is not None else 0) * area 
            if Ptecido == 0:
                msg += "TECIDO: "+tecido+"|"
            peso += Ptecido

            #Calcular peso do tubo
            
            Ptubo = (conn.obterPeso(getTubo(tubo)) if conn.obterPeso(getTubo(tubo)) is not None else 0)     
            Ptubo = Ptubo * lar
            if Ptubo == 0:
                msg += "TUBO: "+tubo+"|"
            peso += Ptubo

            #Calcular peso do perfil            
            Pperfil = conn.obterPeso(getPerfil(perfil))
            Pperfil = (Pperfil if Pperfil is not None else 0) * lar 
            if Pperfil == 0:
                msg += "PERFIL: "+perfil+"|"
            peso += Pperfil
            
            #Calcular o peso do acionamento
            if containsWord(acion, "RS485"):
                Pacion = conn.obterPesobyId(8)
                supMotor = conn.obterPesobyId(4)
                peso += Pacion + supMotor
            elif containsWord(acion, "PVC"):
                Pacion = conn.obterPesobyId(10) * altCom *2 
                peso += Pacion
            elif containsWord(acion, "Metal") or containsWord(acion, "MET") or containsWord(acion, "Met"):
                Pacion = conn.obterPesobyId(11)* altCom * 2
                peso += Pacion
            else:
                Pacion = (conn.obterPeso(acion) if conn.obterPeso(acion) is not None else 0) * altCom * 2 
                if Pacion == 0:
                    msg += "ACION.: "+acion+"|"
                peso += Pacion

        #Atribuindo o peso do volume
        df.loc[i, 'Peso'] = peso
    
    # Adicionando o peso das caixas
    for i in list(df['Volume'].unique()):
        dim = df.loc[df['Volume'] == i, ['Dimensao']]
        dims = dim['Dimensao'].iloc[0].split("x")
        areaV = ((float(dims[0])/1000*4) * (float(dims[2])/1000))
        peso = areaV * conn.obterPesobyId(46)
        dfV = pd.DataFrame(data=[[dim['Dimensao'].iloc[0], i, 'Caixa', peso]], columns=['Nome', 'Volume', 'Tipo', 'Peso'],  index=[-1])
        df = pd.concat([df, dfV], axis=0)

    #Adicionando as observações no banco de dados
    conn.updateObs(msg)

    return df


def getTecido(tecido) -> list:
    #OBJETIVO: dado o nome do tecido, a função retornará a descrição correspondente na tabela.
    listaTecidos = conn.obterOpcoes("tecido")
    listaTecidos = [listaTecidos[i][0] for i in range(0, len(listaTecidos))]
    existe = False
    for i in range(0, len(listaTecidos)):
        if listaTecidos[i] in tecido:
            existe = True
            return listaTecidos[i]
    
    if not existe:
        return ""


def getTubo(tubo) -> list:
    #OBJETIVO: dado o nome do tubo, a função retornará a descrição correspondente na tabela.
    listaTubos = conn.obterOpcoes("tubo")
    listaTubos = [listaTubos[i][0] for i in range(0, len(listaTubos))]
    existe = False

    for i in range(0, len(listaTubos)):
        if listaTubos[i] in tubo:
            existe = True
            return listaTubos[i]

    if not existe:
        return ""


def getPerfil(perfil) -> list:
    #OBJETIVO: dado o nome do tubo, a função retornará a descrição correspondente na tabela.
    listaPerfis = conn.obterOpcoes("perfil")
    listaPerfis = [listaPerfis[i][0] for i in range(0, len(listaPerfis))]
    existe = False

    for i in range(0, len(listaPerfis)):
        if listaPerfis[i] in perfil:
            existe = True
            return listaPerfis[i]

    if not existe:
        return ""


def calcQuantConj(l, componente) -> float:
    #OBJETIVO: Fará o cálculo da quantidade de presilhar que um conjunto vai precisar.
    if containsWord(componente, "PRESILHA"):
        result = 0
        if l <= 1:
            result = 2
        elif l > 1 and l <= 1.5:
            result = 3
        elif l > 1.5 and l <= 2:
            result = 4
        elif l >= 2.5:
            result = 5
        return result
    else: 
        return 1


def calcSupInt(modelo) -> int:
    #OBJETIVO: Fará o cálculo do suporte intermediário quando houver.
    if containsWord(modelo, "Mult Link 2"):
        return 1

    elif containsWord(modelo, "Mult Link 3"):
        return 2

    elif containsWord(modelo, "Mult Link 4"):
        return 3

    elif containsWord(modelo, "Mult Link 5"):
        return 4

    elif containsWord(modelo, "Mult Link 6"):
        return 5

    elif containsWord(modelo, "Mult Link 7"):
        return 6  
    
    else: return 0


