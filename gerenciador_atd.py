import json
import locale
import tkinter as tk
from datetime import datetime, timedelta
from tkinter import messagebox
from typing import Dict

import utilidades

# Definir o locale para português do Brasil, para fazer operações com datas
locale.setlocale(locale.LC_TIME, 'portuguese_brazil')


def main():
    window = tk.Tk()
    min_width = int(window.winfo_screenwidth() * 0.9)
    min_heigth = int(window.winfo_screenheight() * 0.5)
    window.geometry("1480x540")
    window.title("Gerador de ATD")
    window.minsize(min_width, min_heigth)

    def button_executar_onclick():
        messagebox.showinfo(title="Aviso!", message="Tem certeza que deseja gerar o ATD? conferiu os dados?")

        dados_unidade_dict = {"uopm": uopm_var.get(), "uopm_extenso": uopm_extenso_var.get(),
            "grande_comando": grande_comando_var.get(), "uopm_endereco": uopm_endereco_var.get(),
            "uopm_cidade": uopm_cidade_var.get(), "uopm_telefone": uopm_telefone_var.get(),
            "uopm_email": uopm_email_var.get()}

        string_num_portaria = str(num_portaria_var.get()) + "/ATD/CORREG/PMMS/" + data_portaria_var.get()[-4:]
        #
        # dados_atd_dict = {"num_portaria": string_num_portaria, "data_portaria": data_portaria_var.get(),
        #     "nome_encarregado": nome_encarregado_var.get(), "posto_encarregado": posto_encarregado_var.get(),
        #     "mat_encarregado": mat_encarregado_var.get(), "postograd_investigado": postograd_investigado_var.get(),
        #     "nome_investigado": nome_investigado_var.get(), "mat_investigado": mat_investigado_var.get(),
        #     "texto_finalidade": texto_finalidade_var.get(), "nome_escrivao": nome_escrivao_var.get(),
        #     "postograd_escrivao": postograd_escrivao_var.get(), "mat_escrivao": mat_escrivao_var.get(),
        #     "data_autuacaoextenso": data_autuacaoextenso_var.get(), "data_autuacao": string_data_autuacao,
        #     "nome_autinst": nome_autinst_var.get(), "posto_autinst": posto_autinst_var.get(),
        #     "func_autinst": func_autinst_var.get(), "data_relatorio": string_data_relatorio}

        # Salva os dados do da unidade toda vez que executa, pra garantir
        with open("dados_unidade.json", "w", encoding="utf-8") as arquivo_json:
            json.dump(dados_unidade_dict, arquivo_json, ensure_ascii=False, indent=4)

        # Inicialmente é necessário dividir pois os dados da unidade são salvos para futuras criações de IPM
        # mas agora para processarmos todos os documentos é útil juntar as variáveis num objeto só
        # os asteriscos são pra espalhar todos os itens do dicionário, se chamam spread operator
        # dados_gerais: Dict[str, str] = {**dados_unidade_dict, **dados_atd_dict}

        # manipulacoes_atd.processar_capa(dados_gerais)
        # manipulacoes_atd.processar_autuacao(dados_gerais)
        # manipulacoes_atd.processar_portaria_encarregado(dados_gerais)
        # manipulacoes_atd.processar_designacao(dados_gerais)
        # manipulacoes_atd.processar_termo_compromisso(dados_gerais)
        # manipulacoes_atd.processar_despacho(dados_gerais)
        # manipulacoes_atd.processar_recebimento(dados_gerais)
        # manipulacoes_atd.processar_oficio_001(dados_gerais)
        # manipulacoes_atd.processar_oitiva_investigado(dados_gerais)
        # manipulacoes_atd.processar_conclusao(dados_gerais)
        # manipulacoes_atd.processar_relatorio(dados_gerais)
        # manipulacoes_atd.processar_termo_remessa(dados_gerais)
        # manipulacoes_atd.processar_oficio_remessa(dados_gerais)

    # fim função button_executar_onclick

    def confere_dados_unidade():
        try:
            with open("dados_unidade.json", 'r', encoding='utf-8') as f:
                dados = json.load(f)  # Lê o conteúdo do arquivo e converte para um dicionário
                return dados
        except FileNotFoundError:
            print(f"O arquivo dados_unidade.json não foi encontrado.")
            return None
        except json.JSONDecodeError:
            print(f"O arquivo dados_unidade.json não é um JSON válido.")
            return None

    # fim função confere_dados_unidade

    labelgerador = tk.Label(window, text="Gerador de ATD", font="Arial, 24")
    labelgerador.pack(padx=20, pady=20)

    maingrid = tk.Frame(window)
    qtd_colunas = [0, 1, 2, 3, 4, 5]
    for i in qtd_colunas:
        maingrid.columnconfigure(i)

    label_dados_da_unidade = tk.Label(maingrid, text="Dados da Unidade", font="Arial, 16")
    label_dados_da_unidade.grid(row=0, column=0, sticky="w")

    # Linha 1 - UOPM, UOPM por Extenso, Grande Comando
    label_uopm = tk.Label(maingrid, text="UOPM:")
    label_uopm.grid(row=1, column=0, sticky="e")

    uopm_var = tk.StringVar()
    textinput_uopm = tk.Entry(maingrid, textvariable=uopm_var)
    textinput_uopm.grid(row=1, column=1, sticky="w", padx=(0, 25))

    label_uopm_extenso = tk.Label(maingrid, text="UOPM por extenso:")
    label_uopm_extenso.grid(row=1, column=2, sticky="e")

    uopm_extenso_var = tk.StringVar()
    textinput_uopm_extenso = tk.Entry(maingrid, width=50, textvariable=uopm_extenso_var)
    textinput_uopm_extenso.grid(row=1, column=3, sticky="w", padx=(0, 25))

    label_grande_comando = tk.Label(maingrid, text="Grande Comando:")
    label_grande_comando.grid(row=1, column=4, sticky="e")

    grande_comando_var = tk.StringVar()
    textinput_grande_comando = tk.Entry(maingrid, textvariable=grande_comando_var)
    textinput_grande_comando.grid(row=1, column=5, sticky="w", padx=(0, 25))

    # Linha 2 - Endereço da UOPM, Cidade da UOPM, Telefone da UOPM
    label_uopm_endereco = tk.Label(maingrid, text="Endereço da UOPM")
    label_uopm_endereco.grid(row=2, column=0, sticky="e")

    uopm_endereco_var = tk.StringVar()
    textinput_uopm_endereco = tk.Entry(maingrid, width=50, textvariable=uopm_endereco_var)
    textinput_uopm_endereco.grid(row=2, column=1, sticky="w", padx=(0, 25))

    label_uopm_cidade = tk.Label(maingrid, text="Cidade da UOPM")
    label_uopm_cidade.grid(row=2, column=2, sticky="e")

    uopm_cidade_var = tk.StringVar()
    textinput_uopm_cidade = tk.Entry(maingrid, width=50, textvariable=uopm_cidade_var)
    textinput_uopm_cidade.grid(row=2, column=3, sticky="w", padx=(0, 25))

    label_uopm_telefone = tk.Label(maingrid, text="Telefone da UOPM")
    label_uopm_telefone.grid(row=2, column=4, sticky="e")

    uopm_telefone_var = tk.StringVar()
    textinput_uopm_telefone = tk.Entry(maingrid, width=50, textvariable=uopm_telefone_var)
    textinput_uopm_telefone.grid(row=2, column=5, sticky="w", padx=(0, 25))

    # Linha 3 - Email da UOPM
    label_uopm_email = tk.Label(maingrid, text="Email da UOPM")
    label_uopm_email.grid(row=3, column=0, sticky="e")

    uopm_email_var = tk.StringVar()
    textinput_uopm_email = tk.Entry(maingrid, width=50, textvariable=uopm_email_var)
    textinput_uopm_email.grid(row=3, column=1, sticky="w", padx=(0, 25))

    # Linha 4 - Dados do ATD (Título)
    label_dados_do_atd = tk.Label(maingrid, text="Dados do ATD", font="Arial, 16")
    label_dados_do_atd.grid(row=4, column=0, sticky="w")

    # Linha 5 - Número da Portaria, Data da Portaria, Nome do Encarregado
    label_num_portaria = tk.Label(maingrid, text="Número da Portaria (somente número)")
    label_num_portaria.grid(row=5, column=0, sticky="e")

    num_portaria_var = tk.IntVar()
    textinput_num_portaria = tk.Entry(maingrid, width=50, textvariable=num_portaria_var)
    textinput_num_portaria.grid(row=5, column=1, sticky="w", padx=(0, 25))

    label_data_portaria = tk.Label(maingrid, text="Data da Portaria")
    label_data_portaria.grid(row=5, column=2, sticky="e")

    data_portaria_var = tk.StringVar()
    data_portaria_var.set(datetime.today().strftime("%d/%m/%Y"))

    def aplicar_mascara_data_portaria(event):
        texto = data_portaria_var.get()
        # Adiciona "/" após o dia ou mês quando necessário e reposiciona o cursor
        if len(texto) == 2 and texto[1] != "/":
            data_portaria_var.set(texto + "/")
            textinput_data_portaria.icursor(3)  # Coloca o cursor após o "/"
        elif len(texto) == 5 and texto[4] != "/":
            data_portaria_var.set(texto + "/")
            textinput_data_portaria.icursor(6)  # Coloca o cursor após o segundo "/"

    # fim aplicar_mascara_data_portaria

    comando_de_validacao_de_datas = (window.register(utilidades.validar_data_padrao_br), "%P")
    textinput_data_portaria = tk.Entry(maingrid, width=50, textvariable=data_portaria_var, validate="key",
                                       validatecommand=comando_de_validacao_de_datas)
    textinput_data_portaria.grid(row=5, column=3, sticky="w", padx=(0, 25))
    textinput_data_portaria.bind("<KeyRelease>", aplicar_mascara_data_portaria)

    # Chamando função que confere se o arquivo "dados_unidade.json" existe:
    dados_unidade_dict = confere_dados_unidade()
    if dados_unidade_dict is not None:
        # Povoa os entrys corretos com os dados da unidade
        uopm_var.set(dados_unidade_dict["uopm"])
        uopm_extenso_var.set(dados_unidade_dict["uopm_extenso"])
        grande_comando_var.set(dados_unidade_dict["grande_comando"])
        uopm_endereco_var.set(dados_unidade_dict["uopm_endereco"])
        uopm_cidade_var.set(dados_unidade_dict["uopm_cidade"])
        uopm_telefone_var.set(dados_unidade_dict["uopm_telefone"])
        uopm_email_var.set(dados_unidade_dict["uopm_email"])
    # fim do if

    maingrid.pack(side="top")
    window.mainloop()  # Fim do main()