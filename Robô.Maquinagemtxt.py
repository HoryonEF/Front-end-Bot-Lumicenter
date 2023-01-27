import pyautogui as pg
from datetime import datetime
from tkinter import* #força toda a biblioteca
from tkinter import messagebox
from datetime import date, timedelta
import datedelta
import keyboard
import keyboard, time, sys
from threading import Thread
import pandas as pd
import socket as so#comunicação entre processo IPC
from tkinter import filedialog, messagebox
import gspread #documentation https://docs.gspread.org/en/v5.7.0/oauth2.html
from datetime import*
import customtkinter
from tkinter import *
from tkinter import filedialog, messagebox
import time as tm
from datetime import date, timedelta #modulo para adquirir o fuso horário
import datedelta #modulo voltado para realizar calculos com data/hora
from cryptography.fernet import Fernet








#-------------------------------------------------------------------------------------


'''versão com o slider de mudança de velocidade + elementos típicos'''
#versão com o slider de mudança de velocidade + elementos típicos






#------------------------------------------------------	---------------------------------------------------------------
#Indexando data sobre o sistema operacional
windows = so.gethostname()




#criando listas para identificar qual computador está logado e direcionar as respectivas imagens
lista_hosts =  ['LMC00PC158','LMC00PC030','LMC00PC029','LMC00PC219','LMC00PC157','LMC00PC156','LMC00PC248']
lista_names = ["Eduardo.Dalpiaz","Sidnei.Junior","Rafael.Paz","Brener.Gomes","Fernando.Kava","Eduardo.Andrade","Sonias.Computer"]
lista_clicks7 =['PDM-7.PNG','PDM-7.PNG','PDM-7.PNG','PDM-7.PNG','PDM-7.PNG','PDM-7.PNG','PDM-7.PNG'] #imagem imprimir
lista_clicks8 =['PDM-8.PNG','PDM-8.123456.PNG','PDM-8.123456.PNG','PDM-8.123456.PNG','PDM-8.123456.PNG','PDM-8.PNG','PDM-8.123456.PNG'] #ok mandar imprimir
lista_clicks8_1 =['PDM-8.PNG','PDM-8.11.PNG','PDM-8.11.PNG','PDM-8.11.PNG','PDM-8.11.PNG','PDM-8.PNG','PDM-8.11.PNG'] #ok mandar imprimir
lista_clicks11 =['PDM-11.PNG','PDM-11.1.PNG','PDM-11.1.PNG','PDM-11.1.PNG','PDM-11.1.PNG','PDM-11.PNG','PDM-11.1.PNG'] #_pdm_parametro_todas_as_folhas
lista_clicks10 =['PDM-10.PNG','PDM-10.1.PNG','PDM-10.1.PNG','PDM-10.1.PNG','PDM-10.1.PNG','PDM-10.PNG','PDM-11.1.PNG'] #ok confirmar impressão


LMC00PC158 = "Eduardo.Dalpiaz"
LMC00PC030 = "Sidnei.Junior"
LMC00PC029 = "Rafael.Paz"
LMC00PC219 = "brener.gomes"
LMC00PC157 = "Fernando.Kava"
LMC00PC156 = "Eduardo.Andrade"
LMC00PC248 = "Sonias.Computer"




#percorrendo a lista e verificando se há uma correspondência
count = 0

for numero in lista_hosts:
    
    if numero == windows:
    #if numero == LMC00PC248:
        
    
        nome = lista_names.pop(count)
        print('Seja Bem Vindo: '+nome+'\nHostname: '+windows)




        #index posição na lista; variavel será retornada dentro do loop
        nome_lista_clicks8 = lista_clicks8.pop(count)

        
        #index posição na lista; variavel será retornada dentro do loop
        nome_lista_clicks8_1 = lista_clicks8_1.pop(count)

        
        #index posição na lista; variavel será retornada dentro do loop
        nome_lista_clicks11 = lista_clicks11.pop(count)

        #index posição na lista; variavel será retornada dentro do loop
        nome_lista_clicks10 = lista_clicks10.pop(count)
        
    else:
        count += 1





#---------------------------------------------------------------------------------------------------------------

##Indexando informações necessarias##
data_e_hora_atuais = datetime.now()
data_e_hora_em_texto = data_e_hora_atuais.strftime('%d/%m/%Y %H:%M')#modulando a str time
TG = "6. TG".upper()#tornando a str maiuscula
LASER = "7. LASER".upper()
SOBRA = "SOBRA".upper()


#---------------------------------------------------------------------------------------------------------------

###Estrutura UI###

#IMPORTANTE : O método de locação das widgets e das labels foi o metodo Grid


janela =Tk()
janela.title("Robô - "+nome)
janela.geometry("520x388")
janela.configure(bg = "black")

#metodo trava o redimensionamento do widget 
janela.resizable(width=False, height=True)





label = Label(janela, text="Digite o Número de OPs:")
label.grid(row=4, column=1)
label.configure(bg = "black",fg = "white")

#entrada de dados para o número de loops
entrada = Entry(janela, width=20,justify = "center")
entrada.grid(row=4, column=2)










#---------------------------------------------------------------------------------------------------------------

'''Função destinada para realizar a contagem da quantidade de ops do relatorio excel retirado do TOTVS e posteriormente a baixa na lista da google metal mecanica, esclusiva para o setor de maquinagem'''
#Função destinada para realizar a contagem da quantidade de ops do relatorio excel retirado do TOTVS e posteriormente a baixa na lista da google metal mecanica, esclusiva para o setor de maquinagem

def askquestion():
    


    Impressao_Excel = tk.messagebox.askquestion('Impressão .xlsx','Deseja Imprimir pelo Excel?' ,icon = 'warning')#widget de pergunta ao usuario

    if Impressao_Excel == 'yes':


        selecao = filedialog.askopenfilename()#seleção do documento.xlsx

        selecao = selecao.replace('/','\\\\')#troca da barra

        df = pd.read_excel(selecao)
        df = df.iloc[:,8]#filtragem e localização dos dados convenientes no df

        global linhas
        linhas = len(df)
 




        label = Label(janela, text="Digite o Número de OPs:")
        label.grid(row=4, column=1)
        label.configure(bg = "black",fg = "white")


        
        entry_text = tk.StringVar()

        new_text =linhas
        entry_text.set(new_text)

        global entrada
        #entrada de dados para o npumero de loops
        entrada = Entry(janela, width=20,justify = "center",textvariable = entry_text)
        entrada.grid(row=4, column=2)



#---------------------------------------------------------------------------------------------------------------
'''Função destinada para monitorar se os botões foram pressionados'''
#Função destinada para monitorar se os botões foram pressionados        
def botao_win11(event):
    Key = event.char
    if Key =='*':
 
        #e1 = Button(janela,text="Imprimir win 10",command=contagem_regressiva)
        e1 = customtkinter.CTkButton(master=janela, text="Imprimir win 10", command=contagem_regressiva)
        e1.grid(row=14, column=2)
    if Key =='[':

        Thread(target = askquestion).start()
    if Key =='b':

        Thread(target = askquestionn).start()

 
janela.bind('<Key>',botao_win11)



#---------------------------------------------------------------------------------------------------------------
#SLIDER



#função para plotar o slider


# tk basic Scale
var_aux2 = tk.IntVar()
#elementos visuais do slider
scale1 = tk.Scale(janela,bg = "black",fg = "white",cursor = "dot",highlightbackground = "black", from_=0, to=6, orient="horizontal", variable = var_aux2).grid(row = 0, column =2)
medidor = str(var_aux2.get())
var_aux2.set(3)
        



#---------------------------------------------------------------------------------------------------------------

##imagem_FRAME1##

#indexando o Caminho da imagem
imagem = tk.PhotoImage(file = 'lumicenter.png')

#Alocando a imagem na estrutura UI
w = tk.Label(janela, image=imagem)

#Equivalento a imagem a label
w.imagem = imagem
w.grid(row=3, column=2)



        
#---------------------------------------------------------------------------------------------------------------

#Um dicionário em Python é uma coleção com elementos chave-valor que permite representar melhor o mundo real
a = {"value": 0}


'''Função criada para monitorar a tecla space'''
#Função criada para monitorar a tecla space
def monitorKey():
    
    while True:
        if keyboard.is_pressed('space'):


            a['value'] += 1

                     
            break
#função intermediaria focada em iniciar as funções de parada e o bot propiamente dito
        
'''Função criada para interromper o loop e fechar o widget'''
#Função criada para interromper o loop e fechar o widget

def botao_parada():

    while (a['value'] != 0):

        #Condição de parada
        if a['value'] != 0:

            janela.destroy()
            pg.moveTo(x=0, y=0)
            
            break
        break





#---------------------------------------------------------------------------------------------------------------


sec = None

'''Função criada para interromper o loop,contar 4 segundos e encarrgada de fechar o widget'''
#Função criada para interromper o loop,contar 4 segundos e encarrgada de fechar o widget

def tick():

    global sec
    if sec == None:
        sec = int(4)
    if sec == 0:
        time['text'] = 'PROGRAMA INICIADO'
        tm.sleep(1)
        sec = None
        Thread(target = Impressão_Desenho).start()#metodo thread starta o bot impressão desenho
        Thread(target = monitorKey).start()#metodo thread starta o monitoramento de parada
        Thread(target = botao_parada).start()#metodo thread starta o monitoramento de parada

    
    else:
        sec = sec - 1
        time['text'] = sec
        time.after(1000, tick)
        

time = Label(janela, fg='green', bg='black')
time.grid(row=7, column=2)

#---------------------------------------------------------------------------------------------------------------



'''Função criada para realizar a movimentação dos periféricos'''

##COMANDO IMPRIMIR##


def Impressão_Desenho() :#Função criada para realizar a movimentação dos periféricos

    
    
    time['fg'] = 'green'
    n = int(entrada.get())
    previsao = str((2+6+1+1+1)*n + 3)
    m = str(n)
    i=1
    inicio = tm.time()

    var_aux = tk.IntVar()
    bytess = 0
    maxbytes = 0
    
    

    p.grid_remove()
    #p1.grid_forget()



    while (i <= n):
        
        med = int(var_aux2.get())
        medidor = str(med)
        
        
        if medidor == str(6):

            tempo1 = 0.1
            tempo2 = 0.7
            tempo3 = 0.7
            tempo4 = 0.3
            tempo5 = 0.9
            tempo6 = 0.1

        if medidor == str(5):

            
            tempo1 = 0.1
            tempo2 = 1
            tempo3 = 1.2
            tempo4 = 0.8
            tempo5 = 0.5
            tempo6 = 0.6

        if medidor == str(4):


  
            tempo1 = 0.1
            tempo2 = 1.2
            tempo3 = 1.2
            tempo4 = 1.1
            tempo5 = 0.6
            tempo6 = 1

        if medidor == str(3):


            tempo1 = 0.4
            tempo2 = 1.6
            tempo3 = 1.8
            tempo4 = 1.2
            tempo5 = 0.5
            tempo6 = 1

            
        if medidor == str(2):



            tempo1 = 0.5
            tempo2 = 1.7
            tempo3 = 1.9
            tempo4 = 1.5
            tempo5 = 0.6
            tempo6 = 1

        if medidor == str(1):
            

            tempo1 = 0.5
            tempo2 = 1.8
            tempo3 = 2
            tempo4 = 1.6
            tempo5 = 0.8
            tempo6 = 1.5
    
            
        if medidor == str(0):
 
            tempo1 = 0.5
            tempo2 = 2
            tempo3 = 2.3
            tempo4 = 1.8
            tempo5 = 1
            tempo6 = 1.5
      

        #setando a variavel de mudança da progressbar
        var_aux.set(((i-1)*100/(n)))
        #setando as configurações da progressbar
        progress = ttk.Progressbar(
        janela, orient = "horizontal",
        length = 200
        , mode = "determinate",
        variable = var_aux).grid(row=9,column=2)



        #setando a variavel de mudança da progressbar
        var_aux.set(((i-1)*100/(n)))
        #setando as configurações da progressbar
        progress = ttk.Progressbar(
        janela, orient = "horizontal",
        length = 200
        , mode = "determinate",
        variable = var_aux).grid(row=9,column=2)

                 
        pg.hotkey('win', '5')#localização excel
        pg.hotkey('ctrl', 'c')
        pg.press('down')
        pg.hotkey('win', '4')#localização pdm
        if windows == LMC00PC158 or windows == LMC00PC156:
            tm.sleep(tempo1)
            pg.click(885, 195)
            tm.sleep(tempo1)
            pg.doubleClick(885, 195)# ou click na imagem'PDM-6.PNG'
        else:
            tm.sleep(tempo1)
            pg.doubleClick(885, 195)# ou click na imagem'PDM-6.PNG'
            tm.sleep(tempo1)
        
        pg.hotkey('ctrl', 'v')
        pg.press('enter')
        tm.sleep(tempo2)
        tm.sleep(0.1)

        
        

        if pg.locateOnScreen('PDM-7.PNG'):
            try:
                pg.click('PDM-7.PNG')#imagem imprimir
            except:
                tm.sleep(0.1)
                pg.click('PDM-7.PNG')#imagem imprimir
        else:
            msg = tk.messagebox.showerror(title = 'erro',message = 'Não foi possível localizar o icone do PDM')
            break
       



        
        tm.sleep(tempo3)

        
        pg.rightClick(692, 727)
        tm.sleep(tempo4)
        tm.sleep(1)
        if pg.locateOnScreen(nome_lista_clicks8):
            try:
                pg.click(nome_lista_clicks8)#ok mandar imprimir
                tm.sleep(tempo5)
            except:
                msg = tk.messagebox.showerror(title = 'erro',message = 'Não foi localizado uma imagem!')
        else :
            try:
                
                pg.click(nome_lista_clicks8_1)#ok mandar imprimir
                tm.sleep(tempo5)
            except:
                msg = tk.messagebox.showerror(title = 'erro',message = 'Não foi localizado uma imagem!')

                

        tm.sleep(tempo5)


        
        if pg.locateOnScreen(nome_lista_clicks11):
            try:

                pg.click(nome_lista_clicks11)#parametro, todas as folhas
                tm.sleep(1)

                pg.click(nome_lista_clicks10)#ok confirmar impressão
            except:
                pg.click(nome_lista_clicks10)#ok confirmar impressão
                msg = tk.messagebox.showerror(title = 'erro',message = 'Não foi localizado a imagem Todas as folhas!')

        else:
            try:

                pg.click(nome_lista_clicks10)#ok confirmar impressão
            except:
                msg = tk.messagebox.showerror(title = 'erro',message = 'Não foi localizado uma imagem ok mandar imprimir!')


        if pg.locateOnScreen(nome_lista_clicks11) and windows =='LMC00PC158':
        
      
            pg.click(nome_lista_clicks11)#parametro, todas as folhas
            tm.sleep(1)
   
            pg.click(nome_lista_clicks10)#ok confirmar impressão
                
        elif pg.locateOnScreen(nome_lista_clicks10) and windows =='LMC00PC158':

            try:
                pg.click('PDM-10.win10.PNG')#ok confirmar impressão

            except:
                pg.click(nome_lista_clicks10)#ok confirmar impressão








            




        
        tm.sleep(tempo6)

        if i==n:
            
            var_aux.set((i)*100/(n))#linha adicionada para reiniciar a progressbar
            time['text'] = 'PROGRAMA FINALIZADO'
            time['fg'] = 'red'
            fim = tm.time()
            Tempo_de_Operaçao = str(round(fim-inicio,2))
            Tempo_de_Operaçao_minutos = str(round(((fim-inicio)/60),2))
            MsgBox = tk.messagebox.askquestion ('Exit Application','Tempo de Operação: ' + Tempo_de_Operaçao + ' segundos e '+ Tempo_de_Operaçao_minutos + ' minutos\n       '+ m + ' Baixas Realizadas!\n\n\n Deseja Sair da aplicação?' ,icon = 'warning')
            if MsgBox == 'yes':
                janela.destroy()  
                i += 1
                
            else:
                time['fg'] = 'green'
                i = 1            
                break
        else:i+=1
    time['text'] = 'Click em Imprimir Para iniciar novamente'
    time['fg'] = 'YELLOW'
    var_aux.set(((i-1)*100/(n)))
    b = {"value": 0}
          ##              


'''Análogo ao def tick: função criada para interromper o loop,contar 4 segundos e encarrgada de fechar o widget'''
#Análogo ao def tick: função criada para interromper o loop,contar 4 segundos e encarrgada de fechar o widget 
sec = None


def contagem_regressiva():

    global sec
    if sec == None:
        sec = int(4)
    if sec == 0:
        time['text'] = 'PROGRAMA INICIADO'
        tm.sleep(1)
        sec = None
        Thread(target = Impressão_Desenhob).start()#metodo thread starta o bot impressão desenho
        Thread(target = monitorKey).start()#metodo thread starta o monitoramento de parada
        Thread(target = botao_parada).start()#metodo thread starta o monitoramento de parada


    
    else:
        sec = sec - 1
        time['text'] = sec
        time.after(1000, contagem_regressiva)
        

time = Label(janela, fg='green', bg='black')
time.grid(row=7, column=2)





    


#---------------------------------------------------------------------------------------------------------------


'''Análogo ao def Impressão_Desenho:Função criada para realizar a movimentação dos periféricos + baixas no icone do pdm antigo'''

##COMANDO IMPRIMIR##


def Impressão_Desenhob() :#Análogo ao def Impressão_Desenho:Função criada para realizar a movimentação dos periféricos




    time['fg'] = 'green'
    n = int(entrada.get())
    previsao = str((2+6+1+1+1)*n + 3)
    m = str(n)
    i=1
    inicio = tm.time()

    var_aux = tk.IntVar()
    #variavel da barra de progresso
    bytess = 0
    #limite da barra de progresso
    maxbytes = 0
    
    
    #ao iniciar a função eliminamos alguns botoes para deixar o visual limpo
    p.grid_remove()
    p.grid_forget()
    #e1.grid_forget()



    while (i <= n):
        
        med = int(var_aux2.get())
        medidor = str(med)
        
        
        if medidor == str(6):

            tempo1 = 0.1
            tempo2 = 0.82
            tempo3 = 0.85
            tempo4 = 0.3
            tempo5 = 0.99
            tempo6 = 2

        if medidor == str(5):
            
            tempo1 = 0.1
            tempo2 = 1
            tempo3 = 1.2
            tempo4 = 0.8
            tempo5 = 0.5
            tempo6 = 2.5

        if medidor == str(4):
            
            tempo1 = 0.1
            tempo2 = 1.2
            tempo3 = 1.2
            tempo4 = 1.1
            tempo5 = 0.6
            tempo6 = 3

        if medidor == str(3):
            
            tempo1 = 0.1
            tempo2 = 1.6
            tempo3 = 2
            tempo4 = 1.3
            tempo5 = 0.5
            tempo6 = 3

            
        if medidor == str(2):
            
            tempo1 = 0.5
            tempo2 = 1.7
            tempo3 = 1.9
            tempo4 = 1.5
            tempo5 = 0.6
            tempo6 = 1

        if medidor == str(1):
            
            tempo1 = 0.5
            tempo2 = 1.8
            tempo3 = 2
            tempo4 = 1.6
            tempo5 = 0.8
            tempo6 = 3
    
            
        if medidor == str(0):
            
            tempo1 = 0.5
            tempo2 = 2
            tempo3 = 3
            tempo4 = 1.8
            tempo5 = 1
            tempo6 = 3
      

        #setando a variavel de mudança da progressbar
        var_aux.set(((i-1)*100/(n)))
        #setando as configurações da progressbar
        progress = ttk.Progressbar(
        janela, orient = "horizontal",
        length = 200
        , mode = "determinate",
        variable = var_aux).grid(row=9,column=2)



        #setando a variavel de mudança da progressbar
        var_aux.set(((i-1)*100/(n)))
        #setando as configurações da progressbar
        progress = ttk.Progressbar(
        janela, orient = "horizontal",
        length = 200
        , mode = "determinate",
        variable = var_aux).grid(row=9,column=2)   


    #loop ação mecânica 

        pg.hotkey('win', '5')#localização excel
        pg.click(30, 1026)
        pg.hotkey('ctrl', 'c')
        pg.press('down')
        pg.hotkey('win', '4')#localização totvs
        tm.sleep(tempo1)
        pg.mouseDown(button='left', x=885, y=201)
        pg.moveTo(300,194)
        pg.mouseUp(button='left', x=300, y=194)
        tm.sleep(tempo1)
        pg.hotkey('ctrl', 'v')
        pg.press('enter')
        tm.sleep(tempo2)
        if pg.locateOnScreen('PDM-EMBALAGEM.PNG'):
            pg.click('PDM-EMBALAGEM.PNG')#imagem imprimir
        else:
            tk.messagebox.showerror(title = 'ERRO', message = "Não foi possível identificar a imagem do do PDM")
            continue
        tm.sleep(tempo3)
        m.sleep(tempo1)
        pg.rightClick(692, 727)
        tm.sleep(tempo4)
        if pg.locateOnScreen('PDM-8.11.PNG'):
            pg.click('PDM-8.11.PNG')#ok mandar imprimir
            tm.sleep(tempo5)
        else:
            pg.click('PDM-1EMBALAGEM.PNG')#ok mandar imprimir
            tm.sleep(tempo5)
        if pg.locateOnScreen('PDM-11.PNG'):
            pg.doubleClick('PDM-11.PNG')#parametro, todas as folhas
        tm.sleep(tempo6)
        pg.doubleClick('PDM-10.PNG')#ok confirmar impressão
          
            
        i += 1

    tk.messagebox.showinfo(title=None, message="Processo Finalizado!" )





#---------------------------------------------------------------------------------------------------------------

'''Funções referentes ao progressbar'''
#Funções referentes ao progressbar


#indexando variáveis 
var_aux = tk.IntVar()
bytess = 0
maxbytes = 0



'''inicializa read_bytes, com o modulo Thread'''

def start():#inicializa read_bytes

    Thread(target = read_bytes).start()# a função foi forçada para a melhor eficiência da aplicação

    
'''A função read_bytes define os elementos da progress como dimensão, orientação etc...'''
def read_bytes():#elementos da progress dimensão orientação ...
    
    progress = ttk.Progressbar(
    janela, orient = "horizontal",
    length = 150, mode = "determinate",
    variable = var_aux).grid(row=9,column=2)
    #variavel global adicionada para atualizar o valor da progress bar
    global bytess    
    while (bytess < 150):        
            var_aux.set(bytess)
            bytess += 50
            if  bytess == 150:
   
                bytess = 0
                break
                

               ## 




#---------------------------------------------------------------------------------------------------------------


'''A função Atualização é responsável por criar uma nova janela, com objetivo de difundir informações de uso e ou atualizações.'''



def atualizacao():#A função Atualização é responsável por criar uma nova janela, com objetivo de difundir informações de uso e ou atualizações.)
    janela = tk.Toplevel()
    janela.title("Atualização")
    janela.geometry("440x340")
    imagemm = tk.PhotoImage(file = "sequência_barra_de_tarefas.PNG")
    imagemm_op_pendente =tk.PhotoImage(file = "Op_não_encontrada.PNG")
    label_text = tk.Label(janela, text = "Recomendamos a utilização do excel para realizar as baixas\n\nNesta atualização foi adicionado três atalhos  com  as sequintes teclas \n\n\n* -> direcionado para a impresão em pc com apenas uma única tela;\nb -> iniciará a função de baixa por API e \n[ -> que starta a função de contagem do número do OPS")
    label_text.grid(row=1,column = 1)
    label_text2 = tk.Label(janela, text = "\n\nSequência Barra de Tarefas")
    label_text2.grid(row=3,column = 1)  
    label_image = tk.Label(janela, image=imagemm)
    label_image.grid(row=4, column = 1)

    janela.mainloop()




##---------------------------------------------------------------------------------------------------------------
#função de contagem regressiva antes do bot iniciar os comandos de impressão; está contagem é referente ao comando e botão de baixa e impressão apenas; existe outras funções com contador que inicia o loop dos pcs com uma unica tela
'''Análogo ao def tick: função criada para interromper o loop,contar 4 segundos e encarrgada de fechar o widget,
função de contagem regressiva antes do bot iniciar os comandos de impressão; está contagem é referente ao comando e botão de baixa e impressão apenas; existe outras funções com contador que inicia o loop dos pcs com uma unica tela'''
def tickk():
    global sec
    if sec == None:
        sec = int(3)
    if sec == 0:
        time['text'] = 'PROGRAMA INICIADO'
        tm.sleep(1)
        sec = None
        Thread(target = baixa).start()# a função foi forçada para a melhor eficiência da aplicação
        Thread(target = monitorKey).start()
        Thread(target = botao_parada).start()
    else:
        sec = sec - 1
        time['text'] = sec
        time.after(1000, tickk)
        

time = Label(janela, fg='green', bg='black')
time.grid(row=7, column=2)



#---------------------------------------------------------------------------------------------------------------

'''A função é responsável por realizar a baixa na lista do KIT.'''

def baixa():#bloco dedicado a função de baixa no KIT 



    
    n = int(entrada.get())
    previsao = str((2+6+1+1+1)*n + 3)
    m = str(n)
    i=1
    inicio = tm.time()

    var_aux = tk.IntVar()
    bytess = 0
    maxbytes = 0
    

    while (i <= n):

        
        var_aux.set(((i-1)*100/(n)))
        progress = ttk.Progressbar(
        janela, orient = "horizontal",
        length = 200
        , mode = "determinate",
        variable = var_aux).grid(row=9,column=2)

        pg.hotkey('win', '5')#clicar excel
        pg.hotkey('ctrl', 'c')
        pg.press('down')
        pg.hotkey('win', '2')#abrir google
        pg.hotkey('ctrl', 'f')
        tm.sleep(1)
        pg.hotkey('ctrl', 'v')
        tm.sleep(6)
        if pg.locateOnScreen('CRHOME.2.PNG'):
            pg.hotkey('win','5')
            pg.press('up')
            pg.press('left')
            pg.press('f2')
            pg.write('Op não encontrada ->')
            pg.press('enter')
            pg.press('right')
            pg.hotkey('win','5')
            i+=1
            continue
        pg.press('esc')
        tm.sleep(2)
        pg.press('right')
        pg.press('right')
        pg.press('right')
        pg.press('right')
        pg.press('right')
        pg.press('right')
        pg.press('right')
        pg.press('right')
        pg.write(data_e_hora_em_texto)
        pg.press('right')
        pg.press('right')
        tm.sleep(0.5)
        pg.write(selected_month.get())
        tm.sleep(0.5)
        pg.press('enter')
        
   
        time['fg'] = 'red'
        if i==n:
            var_aux.set((i)*100/(n))
            time['text'] = 'PROGRAMA FINALIZADO'
            fim = tm.time()
            Tempo_de_Operaçao = str(round(fim-inicio,2))
            MsgBox = tk.messagebox.askquestion ('Exit Application','Tempo de Operação: ' + Tempo_de_Operaçao + ' segundos\n       '+ m + ' Baixas Realizadas!\n\n\n Deseja Sair da aplicação?' ,icon = 'warning')
            if MsgBox == 'yes':
                janela.destroy()  
                i += 1
                
            else:
                time['fg'] = 'green'
                i = 1            
                break
        else:i+=1




# a estrutura UI foi dividida pois necessita de percorrer algumas funções para adicionar os botões


 
choice = Label(janela, text="Favor Selecionar a Operação:")
choice.configure(bg = "black",fg = "white")
choice.grid(row = 6, column = 1)

# create a combobox
selected_month = tk.StringVar()
selected_cb = ttk.Combobox(janela, textvariable=selected_month)

# get variable, criando uma lista e nomeando ela como values
selected_cb['values'] = [TG, LASER,SOBRA]

# prevent typing a value, atribuindo o tipo de value
selected_cb['state'] = 'readonly'

# place the widget, alocando a widget
selected_cb.grid(row = 6, column = 2)

# bind the selected value changes






#---------------------------------------------------------------------------------------------------------------
data_hora = datetime.now()
data_hora = data_hora.strftime('%d/%m/%Y %H:%M')


#---------------------------------------------------------------------------------------------------------------

'''Função incumbida de realizar a leitura do arquivo direcional (.xlsx), descriptografa e criptografar da chave de acesso a API e a baixa na lista.'''



def askquestionn():#Funções referentes ao progressbar
    #---------------------------------------------------------------------------------------------------------------

    '''Devido a presença de referência singular nos utilizaremos do modulo cryptography do fernet baseado no modelo de chave simetrica ; desenvolvido por Python Cryptographic Authority (PYCA); o modulo foi desenvolvido baseado no processod e cryptografia 128-biAES in CBC mode e An HMAC eith SHA-256'''
    '''key (bytes or str) – A URL-safe base64-encoded 32-byte key. This must be kept secret. Anyone with this key is able to create and read messages.'''


    with open ('\caminhoy', 'rb') as filekey:
        chave = filekey.read()#leitura do arquivo-chave no modo binário autentificação
    print('passei fase 1') 
    fernet = Fernet(chave)
    print('passei fase 2')
    print('\caminho')
    with open('\caminho', 'rb') as arquivo_criptografado:
        criptografado = arquivo_criptografado.read()#leitura do arquivo no modo binário que será descriptografado

    descriptografado = fernet.decrypt(criptografado)#.decode('32 url-safe')
    print('passei fase 3')
    with open('\caminho', 'wb') as arquivo_descriptografado:
        arquivo_descriptografado.write(descriptografado)#criação do arquivo no modo binário do arquivo descryptografado






    #---------------------------------------------------------------------------------------------------------------




    gc = gspread.service_account()


    sh = gc.open("Lista diária_Metal Mecânica")#indexando a spreadsheet

    worksheet = sh.worksheet('LISTA')#indexando a worksheet


    planilha = worksheet.get('B:B')#Resposta em formato de lista
    lista_kit = len(planilha)





    #---------------------------------------------------------------------------------------------------------------
    

    Impressao_Excel = tk.messagebox.askquestion('Baixa.xlsx','Deseja dar baixa pelo Excel?' ,icon = 'warning')

    if Impressao_Excel == 'yes':
    #    Thread(target = leitura_arquivo).start()

        selecao = filedialog.askopenfilename()

        selecao = selecao.replace('/','\\\\')


        
        dff = pd.read_excel(selecao, header=None)
        df2 = dff.iloc[:,1]
        df2 = df2.dropna()#elimina as linhas mas modifica o dtype para float64


        lista_baixas_kit = []

        o = 0
        while (o<lista_kit):
            a = str(planilha[o])
            

            lista_baixas_kit.append(str(a))

            o+=1




        #-----------------------------------------------------------------


        #leitura da planilha excel



        linhas_baixa= str(len(df2))
        linhas_baixa= int(linhas_baixa)




        lista_baixas = []



        p=0

        while (p < linhas_baixa):
            

            dado = str(dff.iloc[p,1])
            lista_baixas.append(str(dado))

            
            p+=1




        set1 = set(lista_baixas)
        set2 = set(lista_baixas_kit)


        '''Função responsável por indexar a linha e a coluna.'''



        def search (lista_baixas_kit, lista_baixas):#Função responsável por indexar a linha e a coluna.

            return [(lista_baixas_kit.index(x)) for x in lista_baixas_kit if lista_baixas in x]



        ops_nao_encontradas = []
        
        i = p
        i = i-1
        ii = i-1 
        while (0<i):
            

            
            row = search(lista_baixas_kit,lista_baixas[i])

            if row != []:
                row2 = row[0] + 1
                worksheet.update_cell(row2,10,data_hora)
                tm.sleep(3)
                worksheet.update_cell(row2,12,selected_month.get())
            else:
                ops_nao_encontradas.append(lista_baixas[i])
                    
            #setando a variavel de mudança da progressbar
            var_aux.set(((ii-i)*100/(ii)))
            #setando as configurações da progressbar
            progress = ttk.Progressbar(
            janela, orient = "horizontal",
            length = 200
            , mode = "determinate",
            variable = var_aux).grid(row=9,column=2)



            #setando a variavel de mudança da progressbar
            var_aux.set(((ii-i)*100/(ii)))
            #setando as configurações da progressbar
            progress = ttk.Progressbar(
            janela, orient = "horizontal",
            length = 200
            , mode = "determinate",
            variable = var_aux).grid(row=9,column=2)

            
            time['text'] = 'CARREGANDO'
            time['fg'] = 'green'
            
            
            
            i-=1
#---------------------------------------------------------------------------------------------------------------
            '''criptografando está secção do código por segurança; foi utilizado um método simétrico de cryptografia com base no modulo fernet'''
            #'''criptografando está secção do código por segurança; foi utilizado um método simétrico de cryptografia com base no modulo fernet'''


        chave = Fernet.generate_key()
        print('gerei a chave')
        with open ('U:\caminho', 'rb') as filekey:
            chave = filekey.read()#leitura do arquivo-chave no modo binário autentificação


        print('li a chave')
    
        fernet = Fernet(chave)

        with open('U:\\producao\\SJP\\Maquinagem\\BAIXAS E IMPORTAÇÕES\\04 - Aplicativo.105\\chave_mestra.key', 'rb') as arquivo:
            conteudo_arquivo = arquivo.read()#leitura do arquivo no modo binário que será criptografado
        print('C:\\Users'+ nome +'\\AppData\\Roaming\\gspread\\service_accountt.txt', 'rb')
        criptografado = fernet.encrypt(conteudo_arquivo)
        print('escrevi no arquivo')
        with open('C:\\Users\\'+ nome +'AppData\\Roaming\\service_accountt.txt','wb') as arquivo_criptografado:
            arquivo_criptografado.write(criptografado)#criação do arquivo no modo binário, já criptografado





#---------------------------------------------------------------------------------------------------------------

        
            
        #setando a variavel de mudança da progressbar
        var_aux.set(((ii)*100/(ii)))
        #setando as configurações da progressbar
        progress = ttk.Progressbar(
        janela, orient = "horizontal",
        length = 200
        , mode = "determinate",
        variable = var_aux).grid(row=9,column=2)
        time['text'] = 'PROGRAMA FINALIZADO'
        time['fg'] = 'yellow'
            



        baixas = str(len(lista_baixas) - len(ops_nao_encontradas) - 1)
        msg_ops_nao_encontradas = tk.messagebox.askquestion('Baixa','Baixas: '+ baixas + '    ' + ' \n Ops não encontradas: '+ str(len(ops_nao_encontradas))  + '\n    Deseja importar para excel?')
        if msg_ops_nao_encontradas == 'yes':
            data = datetime.now()
            data = data.strftime('%d-%m-%y.%H.%M')
            
            list_data = []
            Usuario = []
            for i in ops_nao_encontradas:
                list_data.append(data)
                Usuario.append(nome)
            df = pd.DataFrame(zip(ops_nao_encontradas,list_data,Usuario), columns = ['OP','Data','Usuário'])
            #data = str(data)+".xlsx"
            writer = pd.ExcelWriter("U:\\Producao\\SJP\\Maquinagem\\EXCEL - TOTVS\\OPS FORA DA LISTA\\" + str(data) +'-'+ str(nome) +".xlsx"")
            df.to_excel(writer,'Sheet1')
            writer.save()
            time['text'] = 'XLSX EXPORTADO'

            
                




#icone lumicenter
janela.iconphoto(False, tk.PhotoImage(file='lumicenter.png'))
janela.iconphoto(False, tk.PhotoImage(file='lumicenter.png'))


#botão baixa no kit, utilizado para iniciar o comando do pyautogui de baixa na lista do KIT
button1 = Button(janela,text="Baixa",command=tickk)
button1.grid(row=6, column=3)


#Label com o texto solicitando a digitação do usuario
label = Label(janela, text="Digite o Número de OPs:")
label.grid(row=4, column=3)
label.configure(bg = "black",fg = "black")

label11 = Label(janela, text="Digite o Número de OPs:")
label11.grid(row=4, column=3)
label11.configure(bg = "black",fg = "black")




e = Button(janela,text="Imprimir",command=tick)
e.grid(row=4, column=3)




#labels utilizadas para estruturar melhor as distância do widget
label = Label(janela, text="Digite o Número de OPs:")
label.grid(row=8,rowspan=3, column=2)
label.configure(bg = "black",fg = "black")


#labels utilizadas para estruturar melhor as distância do widget
label1 = Label(janela, text="Digite o Número de OPs:")
label1.grid(row=8,rowspan=3, column=2)
label1.configure(bg = "black",fg = "black")




p = Button(janela,text= "Atualizações", command = atualizacao)
p.grid(row=11, column=3)




#labels utilizadas para estruturar melhor as distância do widget
label1 = Label(janela, text="Digite o Número de OPs:")
label1.grid(row=13,rowspan=4, column=2)
label1.configure(bg = "black",fg = "black")



#labels utilizadas para estruturar melhor as distância do widget
label1 = Label(janela, text="Digite o Número de OPs:")
label1.grid(row=14,rowspan=4, column=2)
label1.configure(bg = "black",fg = "black")






janela.mainloop()
