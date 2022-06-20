from datetime import date
import os
import json

#Arquivo armazenará funções gerais

def containsWord(s, w) -> bool:
    #OBJETIVO: indicará se uma palavra está inserida em uma string

    if type(w) is list:
        for i in range(len(w)):
            if (' ' + w[i] + ' ') in (' ' + s + ' '):
                return True
    else:
        return (' ' + w + ' ') in (' ' + s + ' ')
    


def selectPath(pedido) -> str:
    #OBJETIVO: O método fará a seleção do caminho correto para inserir o relatório.
    # O caminho será relativo à data da criação do relatório.
    
    # Variáveis
    today = date.today()
    mes = today.month
    ano = today.year                                #Ano atual
    mes = selectMonth(mes).upper()                  #Mês atual
    path = ""                                       #Caminho final
    rootPath = ""  #Raíz do caminho
    
    with open('C:\SysOp\Reflexa\ROB\expedicao\src\config.json') as f:
        config = json.load(f)

    rootPath = config['path']

    #Criação do caminho
    path = rootPath + str(ano) + "/" 

    #Verificação da existência do caminho 
    if not os.path.exists(path):
        #Caso não exista, criá-lo
        os.mkdir(path, 0o666)
    path = path + mes + "/"

    if not os.path.exists(path):
        #Caso não exista, criá-lo
        os.mkdir(path, 0o666)

    #Retornar o caminho
    return path + pedido + ".xlsx"


def selectMonth(mes) -> str:
    #OBJETIVO: O método retornará o nome do diretório de acordo com o mês do momento da chamada.
    match mes:
        case 1: return "01-Janeiro"
        case 2: return "02-Fevereiro"
        case 3: return "03-Marco"
        case 4: return "04-Abril"
        case 5: return "05-Maio"
        case 6: return "06-Junho"
        case 7: return "07-Julho"
        case 8: return "08-Agosto"
        case 9: return "09-Setembro"
        case 10: return "10-Outubro"
        case 11: return "11-Novembro"
        case 12: return "12-Dezembro"


def larSuficiente(lar, condDoisJun) -> float:
    #Caso a largura inserida seja igual ou inferior a 0.760, retornar 0.5.
    #Caso contrário, retornar 1.

    if lar == 0:
        return 1
    elif lar <= condDoisJun:
        return 0.5
    else: return 1

def selectComp(modelo) -> str:
    #OBJETIVO: Fará a seleção do componente de cada modelo
    
    prePliRom   = ["P010.0", "P020.0", "P030.0",    #Lista de modelos que possuem
                   "P110.1", "P150.0", "P170.0",    #o componente PLISSADA/ROMANA
                   "P180.1", "RM0100", "RM1500",
                   "RM170", "CE010.0", "CE150.0",
                   "CE170.0", "CE180.1"
        ]  
    comPon      = ["QS80", "QS83", "QS87",          #Lista de modelos que possuem
                   "QS85"]                          #o componente COMANDO/PONTEIRA

    preMidCas   = ["R012", "R011"]                  #Lista de modelos que possuem

    if containsWord(modelo, prePliRom):
        return "PRESILHA PLISSADA/ROMANA - "

    elif containsWord(modelo, comPon):
        comp = "SUP. INST. COMANDO/PONTEIRA"
        if containsWord(modelo, "QS85"):
            return comp + " ACMEDA/FASCIA"
        
        elif containsWord(modelo, "QS80"):
            return comp + " DESTRA"
        
        elif containsWord(modelo, "QS87"):
            return comp + " SOMFY"
        
        elif containsWord(modelo, "Mult Link 2"):
            return comp + " DESTRA"

        elif containsWord(modelo, "Mult Link 3"):
            return comp + " DESTRA"

        elif containsWord(modelo, "Mult Link 4"):
            return comp + " DESTRA"

        elif containsWord(modelo, "Mult Link 5"):
            return comp + " DESTRA"

        elif containsWord(modelo, "Mult Link 6"):
            return comp + " DESTRA"

        elif containsWord(modelo, "Mult Link 7"):
            return comp + " DESTRA"    

        else: 
            return comp + "INDETERMINADO"

    elif containsWord(modelo, preMidCas):
        return "PRESILHA MIDI CASSETE"
    
    elif containsWord(modelo, "QS81"):
        return "PRESILHA MAGNA PLUS"

    elif containsWord(modelo, "QS82"):
        return "PRESILHA MAGNA SILENCIOSA"
    
    elif containsWord(modelo, "R210"):
        return "PRESILHA DIA E NOITE"
    
    elif containsWord(modelo, "R211"):
        return "SUP. INST. ACMEDA"
    
    else: 
        return "CONJ. INDETERMINADO"
    

def cubagem(peso, lar, alt, com) -> float:
    #OBJETIVO: Fará o cálculo da cubagem e retornará o maior valor entre o peso do volume e a cubagem
    # peso : Float
    # lar : Float -> Largura em metros
    # alt : Float -> Altura em metros
    # com : Float -> Comprimento em metros

    cubagem = lar*alt*com*300
    if peso > cubagem:
        return peso
    else:
        return cubagem

if __name__ == "__main__":
    print(cubagem(15, 0.25, 0.27, 1.4))