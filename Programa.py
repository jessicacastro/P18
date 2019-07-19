import tkinter as tk
import tkinter.filedialog as fdlg
import tkinter.scrolledtext as tkst
from tkinter import *
from tkinter.constants import END,HORIZONTAL, VERTICAL, NW, N, E, W, S, SUNKEN, LEFT, RIGHT, TOP, BOTH, YES, NE, X, RAISED, SUNKEN, DISABLED, NORMAL, CENTER
import xlrd
import xlwt
import os
import time
import os.path
import glob
import unicodecsv as csv
class TesteDialogs(object):
    appname= "Divisor de Excel"
    frameWidth      = 800
    frameHeight     = 400
    padx            = 5
    pady            = 5

    flag = 0
    nomeArq = ""
    nomeDir = ""
    pasta = ""
    localArquivo=""
    arrayDelete = []
    colunasGeral = []
    data = ""
    numTotalFiltro = 0
    jaFeito = 0
    def __init__(self, **kw):

        self.root = tk.Tk()

        self.root.title(self.appname)

        self.root.geometry('%dx%d' % (self.frameWidth, self.frameHeight))


        self.msg = Label( text="Selecione o arquivo")
        self.msg["font"] = ("Verdana", "10","bold")
        self.msg.pack (pady= 5)

        self.selecionar = Button()
        self.selecionar["text"] = "Selecionar Arquivo"
        self.selecionar["font"] = ("Calibri", "10")
        self.selecionar["width"] = 20
        self.selecionar["command"] = self. diretorioArquivo
        self.selecionar.pack (pady= 5)

        self.msg = Label( text="Selecione o diretório")
        self.msg["font"] = ("Verdana", "10","bold")
        self.msg.pack (pady= 5)

        self.selecionar = Button()
        self.selecionar["text"] = "Selecionar diretório"
        self.selecionar["font"] = ("Calibri", "10")
        self.selecionar["width"] = 20
        self.selecionar["command"] = self.diretorioFim
        self.selecionar.pack (pady= 5)

        self.msg = Label( text= "Arquivo:")
        self.msg["font"] = ("Verdana", "10","bold")
        self.msg.pack (pady= 10)

        self.msg2 = Label( text= "")
        self.msg2["font"] = ("Verdana", "10","bold")
        self.msg2.pack (pady= 10)

        self.msg = Label( text= "Diretório:")
        self.msg["font"] = ("Verdana", "10","bold")
        self.msg.pack (pady= 10)

        self.msg3 = Label( text= "")
        self.msg3["font"] = ("Verdana", "10","bold")
        self.msg3.pack (pady= 10)

        self.selecionar = Button()
        self.selecionar["text"] = "INICIAR"
        self.selecionar["font"] = ("Calibri", "10")
        self.selecionar["width"] = 30
        self.selecionar["command"] = self.verificar
        self.selecionar.pack ()

        self.minhaTela = tk.Frame(self.root)
        self.minhaTela.pack( padx= "5",expand=1, fill="both")

    def diretorioArquivo(self):
        #primeiro definimos as opções
        opcoes = {}                 # as opções são definidas em um dicionário
        #opcoes['defaultextension'] = '.txt'
        #opcoes['filetypes'] = [('Todos arquivos', '.*'), ('arquivos texto', '.txt')]
        opcoes['filetypes'] = [('Excel', '.xls*')]
        opcoes['initialdir'] = ''    # será o diretório atual
        opcoes['initialfile'] = '' #apresenta todos os arquivos no diretorio
        opcoes['parent'] = self.root
        opcoes['title'] = 'Selecionar arquivo'
        global nomeArq
        global localArquivo
        #retorna o NOME de um arquivo

        nomeArquivo= fdlg.askopenfilename(**opcoes)
        nomeArq = nomeArquivo
        localArquivo = nomeArquivo
        #print(nomeArquivo)
        res = nomeArquivo.split("/")
        res = res[len(res)-1]
        self.msg2.config(text=res)

    def diretorioFim(self):
        #primeiro definimos as opções
        global nomeDir
        opcoes = {}                 # as opções são definidas em um dicionário
        #opcoes['defaultextension'] = '.txt'
        #opcoes['filetypes'] = [('Todos arquivos', '.*'), ('arquivos texto', '.txt')]
        opcoes['initialdir'] = ''    # será o diretório atual
        #opcoes['initialfile'] = '' #apresenta todos os arquivos no diretorio
        opcoes['parent'] = self.root
        opcoes['title'] = 'Selecione o diretório'


        nomeDiretorio= fdlg.askdirectory(**opcoes)
        nomeDir = nomeDiretorio
        self.msg3.config(text=nomeDiretorio)

    def execute(self):
        self.root.mainloop()
    
    def verificar(self):
        global nomeArq
        global nomeDir
        if nomeArq == "" or nomeDir == "":
            print("Em branco")
        else:
            self.inicioPrograma()

    def preencherExcel(self, f, p, contFiltro):
        global colunasGeral
        global data
        global pasta
        global localArquivo
        segundoNome = p.split(".")
        filename = segundoNome[0]+"_" + f + ".xls"
        excluirBATCH = filename.split("BATCH_"+data+"_")

        if contFiltro == 0:
            filename = "BATCH_"+data+"_"+ f + ".xls"
        elif contFiltro == 1:
            filename = "BATCH_"+data+"_"+excluirBATCH[1]
        elif contFiltro == 2:
            filename = "BATCH_"+data+"_"+excluirBATCH[1]

        cabecalho = 3
        cont = 0
        x = 10000
        c = localArquivo  #p
        diretorio= pasta+"/"
        filtroGeral = []
        filtroFuturo = []
        if (contFiltro>0):
            c = diretorio+p
        excel_file = xlwt.Workbook(p)
        sheet = excel_file.add_sheet(f)
        book = xlrd.open_workbook(c)
        sh = book.sheet_by_index(0)
        auxProximo = ""
        diretorio = diretorio+filename

        for rx in range(cabecalho):
            for coluna in range(sh.ncols):
                sheet.write(rx, coluna, sh.cell_value(rowx=rx, colx=coluna))
        pulo = cabecalho
        for rx in range(sh.nrows - cabecalho):
            aux =  sh.cell_value(rowx=rx + pulo, colx=colunasGeral[contFiltro])
            auxTexto =  sh.cell_value(rowx=rx + pulo, colx=7)
            auxMontante =  sh.cell_value(rowx=rx + pulo, colx=4)
            if ((auxMontante == 0) or (auxMontante== "-") or (auxTexto.split(" ")[0]== "Explora++o_Industrial_Uso_de_Rede") or (auxTexto.split(" ")[0]== "REDE_PACOTES")):
                continue
            if(contFiltro<=1):
                auxProximo = sh.cell_value(rowx= rx+ pulo, colx=colunasGeral[contFiltro+1])
            if (aux == f):
                for coluna in range(sh.ncols):      
                    sheet.write(cont+cabecalho, coluna, sh.cell_value(rowx=rx+ pulo, colx=coluna))
                if auxProximo not in filtroFuturo:
                    filtroGeral.append(auxProximo) 
                    filtroFuturo.append(auxProximo)
                cont = cont + 1
        excel_file.save(diretorio)

        self.verificarTamanho([filename,contFiltro+1,filtroFuturo])

    def verificarTamanho(self, f):
        global filtroGeral
        global arrayDelete
        global pasta
        global localArquivo
        global numTotalFiltro
        global jaFeito
        diretorio= pasta+"/"
        c= localArquivo#f[0]
        if f[1] > 0:
            c = diretorio+f[0]
        if f[1] == 1:
            jaFeito = jaFeito + 1
            print("Dividindo ("+str(numTotalFiltro)+"/"+str(jaFeito)+"):" + f[0])
        book = xlrd.open_workbook(c)
        sh = book.sheet_by_index(0)
        if sh.nrows > 1000:
            arrayDelete.append(f[0])
            for aux in f[2]: 
                self.preencherExcel(aux, f[0], f[1])

    def inicioPrograma(self):
        global nomeArq
        global nomeDir
        global pasta
        global arrayDelete
        global colunasGeral
        global data
        global numTotalFiltro
        global jaFeito
        jaFeito = 0
        totalLinhas = 0
        linhasExcluidas =0
        dir_= nomeArq
        print("INICIANDO DIVISÃO")
        colunasGeral = [6, 7, 17]
        print(colunasGeral)
        arrayDelete = []
        filtroAtribuicao = []
        res = dir_.split("/")
        res = res[len(res)-1]
        original = res
        data = original.split("_")
        data = data[2]
        pasta = nomeDir+"/Lote_"+data#+"/XLS"
        if os.path.isdir(pasta):
            pasta = pasta +"/XLS"
        else:
            os.mkdir(pasta)
            pasta = pasta +"/XLS"
            os.mkdir(pasta) 

        book = xlrd.open_workbook(nomeArq)
        sh = book.sheet_by_index(0)
        pulo = 3
        for rx in range(sh.nrows - pulo):
            auxAtribuicao =  sh.cell_value(rowx= rx+ pulo, colx=6)
            if auxAtribuicao not in filtroAtribuicao: 
                    filtroAtribuicao.append(auxAtribuicao) 
        numTotalFiltro=len(filtroAtribuicao)
        self.verificarTamanho([original,0,filtroAtribuicao])
        del(arrayDelete[0])
        for delete in arrayDelete:
            caminho = pasta+"/"+delete
            os.remove(caminho)
        arquivo = glob.glob(pasta+'/*.xls')
        total = len(arquivo)
        pasta2 = nomeDir+"/Lote_"+data+"/CSV"
        if os.path.isdir(pasta2):
            print ("CONVERTENDO EM CSV")
        else:
            os.mkdir(pasta2)
            print ("CONVERTENDO EM CSV")
        print("INICIANDO ...")
        for nomeArq_ in arquivo:
            wb = xlrd.open_workbook(nomeArq_)
            nome = nomeArq_.split(".")
            nome = nome[0].split("\\")
            sh = wb.sheet_by_index(0)
            your_csv_file = open(pasta2+"/"+nome[len(nome)-1]+'.csv', 'wb')
            wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)
            for rownum in range(sh.nrows):
                totalLinhas = totalLinhas + 1
                wr.writerow(sh.row(rownum))
            your_csv_file.close()

        linhasExcluidas = total * 3

        print("TOTAL DE LINHAS: "+str(totalLinhas))
        print("TOTAL DE LINHAS EXCLUIR : "+str(linhasExcluidas))
        print("LINHAS COMPARAR : "+str((totalLinhas-linhasExcluidas)))


def main(args):

	appProc=  TesteDialogs()
	appProc.execute()
	return 0

if __name__ == '__main__':
	import sys
	sys.exit(main(sys.argv))