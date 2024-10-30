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
