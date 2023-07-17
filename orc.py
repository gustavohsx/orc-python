import datetime
import pytesseract
import cv2
import editarPlanilha
import tkinter as tk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo

dados_comp = []
lis = []

def iniciar(filenames):
    arquivos = filenames
    cont = 1
    for i in arquivos:
        
        print(f'Iniciando Imagem {cont}')
        text_area.insert(tk.END, f'Iniciado ORC Imagem {cont}' + "\n") 
        janela.update_idletasks()

        image_name = f'{i}'

        imagem = cv2.imread(image_name, 0)

        _, imagem_bin = cv2.threshold(imagem, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

        caminho = r'C:\Program Files\Tesseract-OCR'
        pytesseract.pytesseract.tesseract_cmd = caminho + r'\tesseract.exe'

        texto = pytesseract.image_to_string(imagem_bin, lang="por")

        texto_usar = texto.split('\n')

        texto_proc = []
        for j in range(len(texto_usar)):
            texto_proc.append(texto_usar[j].split(': '))

        procurarTermos(texto_proc, texto_usar)
        print(f'Terminado Imagem {cont}')
        print('-'*20)
        
        text_area.insert(tk.END, f'Imagem {cont} finalizada' + "\n") 
        janela.update_idletasks()

        cont += 1

    text_area.insert(tk.END, 'Finalizado!' + "\n")
    janela.update_idletasks()


def procurarTermos(texto_proc, texto_usar):
    n_aux = 0
    dic_aux = {}

    for k in range(len(texto_proc)):
        for j in range(len(texto_proc[k])):
            if 'DESTINO' == texto_proc[k][j]:
                aux = f'{texto_proc[k][j+1]} {texto_proc[k][j+2]}'
                dic_aux['DESTINO'] = aux
            elif 'CARREGAMENTO' == texto_proc[k][j]:
                aux = texto_proc[k][j+1]
                dic_aux['CARREGAMENTO'] = aux[:-4]
            elif 'EMITENTE' == texto_proc[k][j]:
                dic_aux['EMITENTE'] = texto_proc[k][j+1]
            elif 'VEICULO' == texto_proc[k][j]:
                dic_aux['VEICULO'] = texto_proc[k][j+1]
            elif 'MOTORISTA' == texto_proc[k][j]:
                aux = texto_proc[k][j+1]
                dic_aux['MOTORISTA'] = aux[:-7]
            elif 'Peso por Rua' in texto_proc[k][j]:
                n_aux = k

    for m in range(len(texto_usar)):
        if 'Rua:' in texto_usar[m]:
            t_aux = texto_usar[m].split(' ')
            dic_aux['Rua'] = t_aux[1]
            break
    
    texto_aux = texto_usar[n_aux].split(' ')

    for l in range(len(texto_aux)):
        if 'produtos' in texto_aux[l]:
            dic_aux['Total da Rua'] = texto_aux[l-1]
        elif 'listados' in texto_aux[l]:
            try:
                aux = texto_aux[l+1].split('i')
                dic_aux['Produtos Listados'] = aux[0]
                print('Tinha itens')
            except:
                dic_aux['Produtos Listados'] = texto_aux[l+1]
        elif 'kg' in texto_aux[l]:
            dic_aux['Peso por Rua'] = texto_aux[l-1]
            cubagem = texto_aux[l+1]
            cubagem = cubagem.replace(',', '.')
            cubagem = float(cubagem)
            cubagem = f'{cubagem:,.4f}'
            cubagem = cubagem.replace('.', ',')
            dic_aux['Cubagem'] = cubagem
    lis.append(dic_aux)

def mostrarResultado():
    for c in range(len(lis)):
        print('\n')
        for i in lis[c]:
            print(f'{i}: {lis[c][i]}')


def salvarDadosPlanilha():

    text_area.insert(tk.END, f'Adicionando a Planilha' + "\n") 
    janela.update_idletasks()

    datahoje = str(datetime.datetime.now().date())
    dia = datahoje[-2:]
    dados = []
    for c in range(len(lis)):
        aux = [dia]
        try:
            aux.append(lis[c]['Rua'])
        except:
            aux.append('-')
        try:
            aux.append(lis[c]['Total da Rua'])
        except:
            aux.append('-')
        try:
            aux.append(lis[c]['Produtos Listados'])
        except:
            aux.append('-')
        try:
            aux.append(lis[c]['Peso por Rua'])
        except:
            aux.append('-')
        try:
            aux.append(lis[c]['Cubagem'])
        except:
            aux.append('-')
        aux.append('1')
        dados.append(aux)
        print('\n')
    try:
        editarPlanilha.salvarPlanilha(dados)
        text_area.insert(tk.END, f'Dados Salvos na Planilha com Sucesso!' + "\n") 
        janela.update_idletasks()
    except Exception as e:
        print(e)
        text_area.insert(tk.END, f'Erro ao Salvar Dados na Planilha' + str(e) + "\n") 
        janela.update_idletasks()


def start(filenames):

    text_area.insert(tk.END, 'Iniciando OCR...' + "\n")
    janela.update_idletasks()

    # valor = entry.get()
    iniciar(filenames)
    mostrarResultado()
    salvarDadosPlanilha()

def abrirArquivos():
    filetypes = ( ('Imagem', '*.jpg'), ('Todos os Arquivos', '*.*'))

    filenames = fd.askopenfilenames(title='Selecionar Arquivos', initialdir='/', filetypes=filetypes)

    showinfo(title='Arquivos Selecionados', message=filenames)

    text_area.delete('1.0', tk.END)
    text_area.insert(tk.END, f'Foram selecionados {len(filenames)} arquivos \n')
    janela.update_idletasks()

    if len(filenames) < 1:
        print('Selecione os arquivos')
        abrirArquivos()
    else:
        start(filenames=filenames)

# Cria a janela principal
janela = tk.Tk()
janela.title('OCR')
width= janela.winfo_screenwidth()  
height= janela.winfo_screenheight()
janela.geometry("%dx%d" % (width, height))

# Cria um botão para mostrar o valor digitado
botao = tk.Button(janela, text="Selecione os arquivos", command=abrirArquivos)
botao.pack()

# Cria a área de texto
text_area = tk.Text(janela)
text_area.pack()

# Inicia o loop principal da janela
janela.mainloop()
