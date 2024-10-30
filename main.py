import tkinter as tk
import gerenciador_ipm
import gerenciador_atd

def abrir_gerenciador_ipm():
    janela_inicial.destroy()
    gerenciador_ipm.main()

def abrir_gerenciador_atd():
    janela_inicial.destroy()
    gerenciador_atd.main()

def centralizar_janela(janela, largura, altura):
    # Obter as dimensões da tela
    largura_tela = janela.winfo_screenwidth()
    altura_tela = janela.winfo_screenheight()

    # Calcular a posição x e y para centralizar a janela
    x = (largura_tela - largura) // 2
    y = (altura_tela - altura) // 2

    # Definir a geometria da janela
    janela.geometry(f"{largura}x{altura}+{x}+{y}")

janela_inicial = tk.Tk()
janela_inicial.title("Bem vindo ao gerador de processos")
largura_janela = 500
altura_janela = 350

# Centralizar a janela
centralizar_janela(janela_inicial, largura_janela, altura_janela)

label1 = tk.Label(janela_inicial, text="Escolha qual processo deseja montar:")

# ATD, Sindicância, IT
# IPM, Conselho de Disciplina, Conselho de Bravura

frame = tk.Frame(janela_inicial)

# Criar e posicionar os botões no grid dentro do Frame
icon_atd = tk.PhotoImage(file=r"icons\atd.png")
button_atd = tk.Button(frame, text="ATD", image=icon_atd, compound="top", command=abrir_gerenciador_atd)
button_atd.grid(row=0, column=0, padx=10, pady=10)

icon_sindicancia = tk.PhotoImage(file="icons/sindicancia.png")
button_sindicancia = tk.Button(frame, text="Sindicância", image=icon_sindicancia, compound="top")
button_sindicancia.grid(row=0, column=1, padx=10, pady=10)

icon_it = tk.PhotoImage(file="icons\it.png")
button_it = tk.Button(frame, text="Inquérito Técnico", image=icon_it, compound="top")
button_it.grid(row=0, column=2, padx=10, pady=10)

icon_ipm = tk.PhotoImage(file="icons\ipm1.png")
button_ipm = tk.Button(frame, text="IPM", image=icon_ipm, compound="top", command=abrir_gerenciador_ipm)
button_ipm.grid(row=1, column=0, padx=10, pady=10)

icon_cons_disciplina = tk.PhotoImage(file="icons\consdisciplina.png")
button_cons_disciplina = tk.Button(frame, text="Conselho de Disciplina",
                                   image=icon_cons_disciplina, compound="top")
button_cons_disciplina.grid(row=1, column=1, padx=10, pady=10)

icon_cons_especial = tk.PhotoImage(file="icons\consespecial.png")
button_cons_especial = tk.Button(frame, text="Conselho Especial", image=icon_cons_especial, compound="top")
button_cons_especial.grid(row=1, column=2, padx=10, pady=10)

label1.pack(pady=20)
frame.pack(pady=20)
janela_inicial.mainloop()