import json
import locale
import tkinter as tk
from datetime import datetime, timedelta
from tkinter import messagebox
from typing import Dict
import ctypes

import manipulacoes_ipm
import utilidades

# Definir o locale para português do Brasil, para fazer operações com datas
locale.setlocale(locale.LC_TIME, 'portuguese_brazil')


def main():
    window = tk.Tk()
    largura = int(window.winfo_screenwidth() * 0.9)
    altura = int(window.winfo_screenheight() * 0.7)
    x_pos = 5
    y_pos = 5
    window.geometry(f"{largura}x{altura}+{x_pos}+{y_pos}")
    window.title("Gerador de IPM")
    window.minsize(largura, altura)

    def button_executar_onclick():
        if(not messagebox.askokcancel(title="Aviso!", message="Tem certeza que deseja gerar o IPM? conferiu os dados?")):
            return

        # Pega a data de autuação, informa que é uma string no formato dd/mm/aaaa e cria um objeto Datetime
        data = datetime.strptime(data_autuacao_var.get(), "%d/%m/%Y")
        # Adiciona trinta dias
        nova_data = data + timedelta(days=30)
        # Salva a string da data relatorio calculada
        data_relatorio = nova_data.strftime("%d/%m/%Y")

        dados_unidade_dict = {"uopm": uopm_var.get(), "uopm_extenso": uopm_extenso_var.get(),
            "grande_comando": grande_comando_var.get(), "uopm_endereco": uopm_endereco_var.get(),
            "uopm_cidade": uopm_cidade_var.get(), "uopm_telefone": uopm_telefone_var.get(),
            "uopm_email": uopm_email_var.get()}

        string_num_portaria = str(num_portaria_var.get()) + "/IPM/CORREG/PMMS/" + data_portaria_var.get()[-4:]
        string_data_portaria = utilidades.get_data_semi_extenso(data_portaria_var.get())
        string_data_autuacao = utilidades.get_data_semi_extenso(data_autuacao_var.get())
        string_data_relatorio = utilidades.get_data_semi_extenso(data_relatorio)

        dados_ipm_dict = {"num_portaria": string_num_portaria, "data_portaria": string_data_portaria,
            "nome_encarregado": nome_encarregado_var.get(), "posto_encarregado": posto_encarregado_var.get(),
            "mat_encarregado": mat_encarregado_var.get(), "postograd_investigado": postograd_investigado_var.get(),
            "nome_investigado": nome_investigado_var.get(), "mat_investigado": mat_investigado_var.get(),
            "texto_finalidade": textinput_texto_finalidade.get("1.0", tk.END), "nome_escrivao": nome_escrivao_var.get(),
            "postograd_escrivao": postograd_escrivao_var.get(), "mat_escrivao": mat_escrivao_var.get(),
            "data_autuacaoextenso": data_autuacaoextenso_var.get(), "data_autuacao": string_data_autuacao,
            "nome_autinst": nome_autinst_var.get(), "posto_autinst": posto_autinst_var.get(),
            "func_autinst": func_autinst_var.get(), "data_relatorio": string_data_relatorio}

        # Salva os dados do da unidade toda vez que executa, pra garantir
        with open("dados_unidade.json", "w", encoding="utf-8") as arquivo_json:
            json.dump(dados_unidade_dict, arquivo_json, ensure_ascii=False, indent=4)

        # Inicialmente é necessário dividir pois os dados da unidade são salvos para futuras criações de IPM
        # mas agora para processarmos todos os documentos é útil juntar as variáveis num objeto só
        # os asteriscos são pra espalhar todos os itens do dicionário, se chamam spread operator
        dados_gerais: Dict[str, str] = {**dados_unidade_dict, **dados_ipm_dict}


        try:
            manipulacoes_ipm.processar_capa(dados_gerais)
            manipulacoes_ipm.processar_autuacao(dados_gerais)
            manipulacoes_ipm.processar_portaria_encarregado(dados_gerais)
            manipulacoes_ipm.processar_designacao(dados_gerais)
            manipulacoes_ipm.processar_termo_compromisso(dados_gerais)
            manipulacoes_ipm.processar_despacho(dados_gerais)
            manipulacoes_ipm.processar_recebimento(dados_gerais)
            manipulacoes_ipm.processar_oficio_001(dados_gerais)
            manipulacoes_ipm.processar_oitiva_investigado(dados_gerais)
            manipulacoes_ipm.processar_conclusao(dados_gerais)
            manipulacoes_ipm.processar_relatorio(dados_gerais)
            manipulacoes_ipm.processar_termo_remessa(dados_gerais)
            manipulacoes_ipm.processar_oficio_remessa(dados_gerais)
            manipulacoes_ipm.processar_solucao(dados_gerais)
        except Exception as error:
            messagebox.showerror(title="Erro", message=f"Erro ao criar IPM: \n{error}")
        else:
            nome_da_pasta_de_saida = manipulacoes_ipm.criar_pasta_saida(dados_gerais["num_portaria"])
            messagebox.showinfo(icon=messagebox.WARNING, title="Sucesso", message=f"IPM Criado com sucesso! Confira a pasta {nome_da_pasta_de_saida}")

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

    labelgerador = tk.Label(window, text="Gerador de IPM", font="Arial, 24")
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

    # Linha 4 - Dados do IPM (Título)
    label_dados_do_ipm = tk.Label(maingrid, text="Dados do IPM", font="Arial, 16")
    label_dados_do_ipm.grid(row=4, column=0, sticky="w")

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

        # Captura a posição atual do cursor
        pos_cursor = textinput_data_portaria.index(tk.INSERT)

        # Remove qualquer caractere que não seja número ou "/"
        texto = "".join(c for c in texto if c.isdigit() or c == "/")

        # Garante que a string não ultrapasse 10 caracteres
        if len(texto) > 10:
            texto = texto[:10]

        novo_texto = ""
        for i, char in enumerate(texto):
            if i in [2, 5]:  # Posiciona automaticamente as barras "/"
                if char != "/":
                    novo_texto += "/"
                    pos_cursor += 1  # Ajusta a posição do cursor
            novo_texto += char

        # Atualiza o campo
        data_portaria_var.set(novo_texto)

        # Reposiciona o cursor corretamente
        textinput_data_portaria.icursor(pos_cursor)

    # fim aplicar_mascara_data_portaria

    comando_de_validacao_de_datas = (window.register(utilidades.validar_data_padrao_br), "%P")
    textinput_data_portaria = tk.Entry(maingrid, width=50, textvariable=data_portaria_var, validate="key",
                                       validatecommand=comando_de_validacao_de_datas)
    textinput_data_portaria.grid(row=5, column=3, sticky="w", padx=(0, 25))
    textinput_data_portaria.bind("<KeyRelease>", aplicar_mascara_data_portaria)

    # Linha 6 - Data autuação, Data autuação por extenso
    def aplicar_mascara_data_autuacao(event):
        texto = data_autuacao_var.get()

        # Captura a posição atual do cursor
        pos_cursor = textinput_data_autuacao.index(tk.INSERT)

        # Remove qualquer caractere que não seja número ou "/"
        texto = "".join(c for c in texto if c.isdigit() or c == "/")

        # Garante que a string não ultrapasse 10 caracteres
        if len(texto) > 10:
            texto = texto[:10]

        novo_texto = ""
        for i, char in enumerate(texto):
            if i in [2, 5]:  # Posiciona automaticamente as barras "/"
                if char != "/":
                    novo_texto += "/"
                    pos_cursor += 1  # Ajusta a posição do cursor
            novo_texto += char

        # Atualiza o campo
        data_autuacao_var.set(novo_texto)

        # Reposiciona o cursor corretamente
        textinput_data_autuacao.icursor(pos_cursor)

    # fim aplicar_mascara_data_autuacao

    def atualiza_dataautuacao_extenso(event):
        data_por_extenso = utilidades.get_data_extenso(data_autuacao_var.get())
        data_autuacaoextenso_var.set(data_por_extenso)
    # fim atualiza_dataautuacao_extenso

    label_data_autuacao = tk.Label(maingrid, text="Data de Autuação")
    label_data_autuacao.grid(row=6, column=0, sticky="e")

    data_autuacao_var = tk.StringVar()
    data_autuacao_var.set(datetime.today().strftime("%d/%m/%Y"))
    textinput_data_autuacao = tk.Entry(maingrid, width=50, textvariable=data_autuacao_var, validate="key",
                                       validatecommand=comando_de_validacao_de_datas)
    textinput_data_autuacao.grid(row=6, column=1, sticky="w", padx=(0, 25))
    textinput_data_autuacao.bind("<KeyRelease>", aplicar_mascara_data_autuacao)
    textinput_data_autuacao.bind("<FocusOut>", atualiza_dataautuacao_extenso)

    label_data_autuacaoextenso = tk.Label(maingrid, text="Data aut extenso, tudo minúsculo")
    label_data_autuacaoextenso.grid(row=6, column=2, sticky="e")

    data_autuacaoextenso_var = tk.StringVar()
    sugestao_data_autuacao_extenso = utilidades.get_data_extenso(data_autuacao_var.get())
    data_autuacaoextenso_var.set(sugestao_data_autuacao_extenso)
    textinput_data_autuacaoextenso = tk.Entry(maingrid, width=50, textvariable=data_autuacaoextenso_var)
    textinput_data_autuacaoextenso.grid(row=6, column=3, sticky="w", padx=(0, 25))

    # Linha 7 - Dados do Encarregado
    label_dados_encarregado = tk.Label(maingrid, text="Dados do Encarregado", font="Arial, 16")
    label_dados_encarregado.grid(row=7, column=0, sticky="w", columnspan=6)

    label_nome_encarregado = tk.Label(maingrid, text="Nome do Encarregado")
    label_nome_encarregado.grid(row=8, column=0, sticky="e")

    nome_encarregado_var = tk.StringVar()
    textinput_nome_encarregado = tk.Entry(maingrid, width=50, textvariable=nome_encarregado_var)
    textinput_nome_encarregado.grid(row=8, column=1, sticky="w", padx=(0, 25))

    label_posto_encarregado = tk.Label(maingrid, text="Posto do Encarregado")
    label_posto_encarregado.grid(row=8, column=2, sticky="e")

    posto_encarregado_var = tk.StringVar()
    textinput_posto_encarregado = tk.Entry(maingrid, width=50, textvariable=posto_encarregado_var)
    textinput_posto_encarregado.grid(row=8, column=3, sticky="w", padx=(0, 25))

    label_mat_encarregado = tk.Label(maingrid, text="Matrícula do Encarregado")
    label_mat_encarregado.grid(row=8, column=4, sticky="e")

    def validar_matricula(mat_string):
        # Permite apenas dígitos, e vazio
        if mat_string == "" or mat_string.isdigit():
            return True
        return False

    mat_encarregado_var = tk.StringVar()
    comando_de_validacao_de_matriculas = (window.register(validar_matricula), "%P")
    textinput_mat_encarregado = tk.Entry(maingrid, width=50, textvariable=mat_encarregado_var, validate="key",
                                         validatecommand=comando_de_validacao_de_matriculas)
    textinput_mat_encarregado.grid(row=8, column=5, sticky="w", padx=(0, 25))

    # Linha 9 - Dados do Escrivão
    label_dados_escrivao = tk.Label(maingrid, text="Dados do Escrivão", font="Arial, 16")
    label_dados_escrivao.grid(row=9, column=0, sticky="w", columnspan=6)

    label_nome_escrivao = tk.Label(maingrid, text="Nome do Escrivão")
    label_nome_escrivao.grid(row=10, column=0, sticky="e")

    nome_escrivao_var = tk.StringVar()
    textinput_nome_escrivao = tk.Entry(maingrid, width=50, textvariable=nome_escrivao_var)
    textinput_nome_escrivao.grid(row=10, column=1, sticky="w", padx=(0, 25))

    label_postograd_escrivao = tk.Label(maingrid, text="Posto/Graduação do Escrivão")
    label_postograd_escrivao.grid(row=10, column=2, sticky="e")

    postograd_escrivao_var = tk.StringVar()
    textinput_postograd_escrivao = tk.Entry(maingrid, width=50, textvariable=postograd_escrivao_var)
    textinput_postograd_escrivao.grid(row=10, column=3, sticky="w", padx=(0, 25))

    label_mat_escrivao = tk.Label(maingrid, text="Matrícula do Escrivão")
    label_mat_escrivao.grid(row=10, column=4, sticky="e")

    mat_escrivao_var = tk.StringVar()
    textinput_mat_escrivao = tk.Entry(maingrid, width=50, textvariable=mat_escrivao_var, validate="key",
                                      validatecommand=comando_de_validacao_de_matriculas)
    textinput_mat_escrivao.grid(row=10, column=5, sticky="w", padx=(0, 25))

    # Linha 11 - Dados do Investigado
    label_dados_investigado = tk.Label(maingrid, text="Dados do Investigado", font="Arial, 16")
    label_dados_investigado.grid(row=11, column=0, sticky="w", columnspan=6)

    label_nome_investigado = tk.Label(maingrid, text="Nome do Investigado")
    label_nome_investigado.grid(row=12, column=0, sticky="e")

    nome_investigado_var = tk.StringVar()
    textinput_nome_investigado = tk.Entry(maingrid, width=50, textvariable=nome_investigado_var)
    textinput_nome_investigado.grid(row=12, column=1, sticky="w", padx=(0, 25))

    label_postograd_investigado = tk.Label(maingrid, text="Posto/Graduação do Investigado")
    label_postograd_investigado.grid(row=12, column=2, sticky="e")

    postograd_investigado_var = tk.StringVar()
    textinput_postograd_investigado = tk.Entry(maingrid, width=50, textvariable=postograd_investigado_var)
    textinput_postograd_investigado.grid(row=12, column=3, sticky="w", padx=(0, 25))

    label_mat_investigado = tk.Label(maingrid, text="Matrícula do Investigado")
    label_mat_investigado.grid(row=12, column=4, sticky="e")

    mat_investigado_var = tk.StringVar()
    textinput_mat_investigado = tk.Entry(maingrid, width=50, textvariable=mat_investigado_var, validate="key",
                                         validatecommand=comando_de_validacao_de_matriculas)
    textinput_mat_investigado.grid(row=12, column=5, sticky="w", padx=(0, 25))

    # Linha 13 - Dados da Autoridade Instauradora
    label_dados_autinst = tk.Label(maingrid, text="Dados da Autoridade Instauradora", font="Arial, 16")
    label_dados_autinst.grid(row=13, column=0, sticky="w", columnspan=6)

    label_nome_autinst = tk.Label(maingrid, text="Nome da Autoridade Instauradora")
    label_nome_autinst.grid(row=14, column=0, sticky="e")

    nome_autinst_var = tk.StringVar()
    textinput_nome_autinst = tk.Entry(maingrid, width=50, textvariable=nome_autinst_var)
    textinput_nome_autinst.grid(row=14, column=1, sticky="w", padx=(0, 25))

    label_posto_autinst = tk.Label(maingrid, text="Posto:")
    label_posto_autinst.grid(row=14, column=2, sticky="e")

    posto_autinst_var = tk.StringVar()
    textinput_posto_autinst = tk.Entry(maingrid, width=50, textvariable=posto_autinst_var)
    textinput_posto_autinst.grid(row=14, column=3, sticky="w", padx=(0, 25))

    label_func_autinst = tk.Label(maingrid, text="Função:")
    label_func_autinst.grid(row=14, column=4, sticky="e")

    func_autinst_var = tk.StringVar()
    textinput_func_autinst = tk.Entry(maingrid, width=50, textvariable=func_autinst_var)
    textinput_func_autinst.grid(row=14, column=5, sticky="w", padx=(0, 25))

    # Linha 15 - Texto finalidade
    label_texto_finalidade = tk.Label(maingrid, text="Texto da Finalidade")
    label_texto_finalidade.grid(row=15, column=0, sticky="e")

    # Configurar o peso das colunas (para redimensionar proporcionalmente)
    for col in range(8):  # Número de colunas
        maingrid.columnconfigure(col, weight=1)
    maingrid.rowconfigure(15, weight=1)  # Permitir expansão para a linha do Text

    textinput_texto_finalidade = tk.Text(maingrid, wrap="word")
    textinput_texto_finalidade.grid(row=15, column=1, columnspan=4, sticky="ew", padx=10, pady=10)
    # Adicionar barra de rolagem
    scrollbar = tk.Scrollbar(maingrid, command=textinput_texto_finalidade.yview)
    scrollbar.grid(row=9, column=5, sticky="nsw")  # Posicionada ao lado do `Text`
    textinput_texto_finalidade.configure(yscrollcommand=scrollbar.set)

    # Linha 16 - botão executar GERAR IPM
    button_executar = tk.Button(maingrid, text="Gerar IPM", command=button_executar_onclick)
    button_executar.grid(row=16, column=0, columnspan=2, sticky="nsew", padx=10, pady=10)

    def button_sugerir_onclick():
        # Povoando temporariamente Entrys para fins de teste
        num_portaria_var.set("101")
        nome_encarregado_var.set("William Scaramuzzi Teixeira")
        posto_encarregado_var.set("Major QOPM")
        mat_encarregado_var.set("101826021")
        nome_escrivao_var.set("João da Silva")
        postograd_escrivao_var.set("2º Sargento QOPM")
        mat_escrivao_var.set("011111021")
        nome_investigado_var.set("José da Silva")
        postograd_investigado_var.set("Soldado QPPM")
        mat_investigado_var.set("022222021")
        nome_autinst_var.set("William Scaramuzzi Teixeira")
        posto_autinst_var.set("Major QOPM")
        func_autinst_var.set("Comandante da 14ª CIPM")
        textinput_texto_finalidade.insert("1.0",
                                          "Apurar possíveis irregularidades atribuídas ao investigado quando em abordagem a veículo no dia 05 de março de 2024")

    # Linha 16 - botão sugerir dados para fins de teste
    button_sugerir = tk.Button(maingrid, text="Sugerir dados", command=button_sugerir_onclick)
    button_sugerir.grid(row=16, column=2, columnspan=2)

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


# se esse arquivo .py for rodado como principal, chama a função main()
if __name__ == "__main__":
    main()
