import customtkinter
from tkinter import *

#tema da janela
customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("dark-blue")

#abrindo a janela
janela = customtkinter.CTk()
janela.geometry("700x400")
janela.title("Bot Lumicenter")
janela.resizable(False, False)

def slider_callback(value):
    progressbar_1.set(value)

#inicio do frame
frame_1 = customtkinter.CTkFrame(master=janela)
frame_1.pack(pady=20, padx=60, fill="both", expand=True)

#logo lumicenter
img =  PhotoImage(file="logo.png")
Label_img = customtkinter.CTkLabel(master=frame_1, image=img, text="")
Label_img.pack(pady=10, padx=10)

#texto velocidade
label_1 = customtkinter.CTkLabel(master=frame_1, text="Velocidade", text_color="white", font=("Arial", 12))
label_1.place(x=100, y=18)

#barra progresso
progressbar_1 = customtkinter.CTkProgressBar(master=frame_1)
progressbar_1.place(x = 25, y = 10)

slider_1 = customtkinter.CTkSlider(master=frame_1, command=slider_callback, from_=0, to=1)
slider_1.place(x = 25, y = 40)
slider_1.set(0.5)

#botão Atualizações
Bt3 = customtkinter.CTkButton(master=frame_1, text="Atualizações", width=100)
Bt3.place(x=410, y=10)

#botão baixa
Bt2 = customtkinter.CTkButton(master=frame_1, text="Baixa", width=100)
Bt2.place(x=410, y=50)

#campo número de OPs
entry1 = customtkinter.CTkEntry(master=frame_1, placeholder_text="Digite o número de OPs", width=300, font=("Arial", 14))
entry1.place(x=140, y=140)

#campo Seleção de operação
menu_op = customtkinter.CTkOptionMenu(master=frame_1, width=300 ,values=["Opção1", "Opção2","Opção3"])
menu_op.place(x=140, y=180)
menu_op.set("Selecione a Operação")

#botão Imprimir
Bt1 = customtkinter.CTkButton(master=frame_1, text="Imprimir", width=200)#V
Bt1.place(x=190, y=220)

janela.mainloop()