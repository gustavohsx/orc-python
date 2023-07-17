import openpyxl

def salvarPlanilha(dado):
    arquivo_excel = 'Separação noturno diaria JUNHO 2023 - Copia.xlsm'
    workbook = openpyxl.load_workbook(arquivo_excel)
    nome_planilha = '1'
    planilha = workbook[nome_planilha]

    letras = {1:'A', 2:'B', 3:'C', 4:'D', 5:'E', 6:'F', 7:'G', 8:'H'}

    
    dados = dado
    for i in range(len(dados)):
        cont = 0
        celula = planilha[f'A7']
        linha = 7
        while cont < len(dados[i]):
            if celula.value is None:
                print(f'[{letras[cont+1]}{linha}] Esta vazia')
                planilha.cell(row=linha, column=cont+1, value=dados[i][cont]).value
                print(f'Adicionando {dados[i][cont]} a celula [{letras[cont+1]}{linha}]')
                cont += 1
                celula = planilha[f'{letras[cont+1]}{linha}']
            else:
                print(f'[{letras[cont+1]}{linha}] Nao esta vazia')
                linha += 1
                celula = planilha[f'{letras[cont+1]}{linha}']
    workbook.save(arquivo_excel)
