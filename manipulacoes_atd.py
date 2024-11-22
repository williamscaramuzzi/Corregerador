import os
import re
from typing import Dict
from docxtpl import DocxTemplate


def criar_pasta_saida(num_portaria_completo: str) -> str:
    # Usando expressão regular para capturar os números iniciais do num_portaria_completo
    resultado = re.match(r'(\d+)/ATD/CORREG/PMMS/', num_portaria_completo)
    num = 0
    if resultado:
        num = resultado.group(1)  # Retorna apenas os números capturados

    # Caminho do arquivo main.py
    caminho_principal = os.path.dirname(os.path.abspath(__file__))

    # Nome da nova pasta, num_portaria_completo[-4:] pega o ano 2024 que fica no final de /ATD/CORREG/PMMS/2024
    nome_pasta = "ATD Portaria n " + str(num) + " " + num_portaria_completo[-4:]

    # Caminho completo para a nova pasta
    caminho_pasta = os.path.join(caminho_principal, nome_pasta)

    # Criação da pasta
    if not os.path.exists(caminho_pasta):
        os.makedirs(caminho_pasta)
        print(f"Pasta '{nome_pasta}' criada com sucesso.")
    else:
        print(f"A pasta '{nome_pasta}' já existe.")
    return nome_pasta  # Retorna o nome da pasta independente dela já existir

def processar_capa(obj: Dict[str, str]) -> None:
    documento = DocxTemplate("templates/atd/A - Capa.docx")
    dados = {
        "grande_comando": obj["grande_comando"],
        "uopm_extenso": obj["uopm_extenso"],
        "uopm": obj["uopm"],
        "uopm_cidade": obj["uopm_cidade"],
        "uopm_email": obj["uopm_email"],
        "uopm_endereco": obj["uopm_endereco"],
        "uopm_telefone": obj["uopm_telefone"],
        "num_portaria": obj["num_portaria"],
        "data_portaria": obj["data_portaria"],
        "postograd_acusado": obj["postograd_acusado"],
        "nome_acusado": obj["nome_acusado"],
        "mat_acusado": obj["mat_acusado"],
        "posto_encarregado": obj["posto_encarregado"],
        "nome_encarregado": obj["nome_encarregado"],
        "mat_encarregado": obj["mat_encarregado"],
        "texto_finalidade": obj["texto_finalidade"]
    }
    documento.render(dados)
    nome_da_pasta = criar_pasta_saida(obj["num_portaria"])
    caminho_saida = nome_da_pasta + r"/A - Capa.docx"
    documento.save(caminho_saida)
    return None

def processar_formulario(obj: Dict[str, str]) -> None:
    documento = DocxTemplate("templates/atd/B - Formulario.docx")
    dados = {
        "grande_comando": obj["grande_comando"],
        "uopm_extenso": obj["uopm_extenso"],
        "uopm": obj["uopm"],
        "uopm_cidade": obj["uopm_cidade"],
        "uopm_email": obj["uopm_email"],
        "uopm_endereco": obj["uopm_endereco"],
        "uopm_telefone": obj["uopm_telefone"],
        "num_portaria": obj["num_portaria"],
        "data_portaria": obj["data_portaria"],
        "postograd_acusado": obj["postograd_acusado"],
        "nome_acusado": obj["nome_acusado"],
        "mat_acusado": obj["mat_acusado"],
        "posto_encarregado": obj["posto_encarregado"],
        "nome_encarregado": obj["nome_encarregado"],
        "mat_encarregado": obj["mat_encarregado"],
        "texto_finalidade": obj["texto_finalidade"]
    }
    documento.render(dados)
    nome_da_pasta = criar_pasta_saida(obj["num_portaria"])
    caminho_saida = nome_da_pasta + r"/B - Formulario.docx"
    documento.save(caminho_saida)
    return None

def processar_relatorio(obj: Dict[str, str]) -> None:
    documento = DocxTemplate("templates/atd/Z - Relatorio.docx")
    dados = {
        "uopm": obj["uopm"],
        "grande_comando": obj["grande_comando"],
        "uopm_extenso": obj["uopm_extenso"],
        "uopm_email": obj["uopm_email"],
        "uopm_telefone": obj["uopm_telefone"],
        "uopm_cidade": obj["uopm_cidade"],
        "uopm_endereco": obj["uopm_endereco"],
        "func_autinst": obj["func_autinst"],
        "posto_autinst": obj["posto_autinst"],
        "nome_autinst": obj["nome_autinst"],
        "num_portaria": obj["num_portaria"],
        "data_portaria": obj["data_portaria"],
        "texto_finalidade": obj["texto_finalidade"],
        "postograd_acusado": obj["postograd_acusado"],
        "nome_acusado": obj["nome_acusado"],
        "mat_acusado": obj["mat_acusado"],
        "data_relatorio": obj["data_relatorio"],
        "nome_encarregado": obj["nome_encarregado"],
        "posto_encarregado": obj["posto_encarregado"],
        "mat_encarregado": obj["mat_encarregado"]
    }
    documento.render(dados)
    nome_da_pasta = criar_pasta_saida(obj["num_portaria"])
    caminho_saida = nome_da_pasta + r"/W - Relatorio.docx"
    documento.save(caminho_saida)
    return None