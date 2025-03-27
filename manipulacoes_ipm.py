import os
import re
import sys
from typing import Dict
from docxtpl import DocxTemplate

def get_caminho_templates(caminho):
    """ Retorna o caminho correto da pasta templates """
    if getattr(sys, 'frozen', False):  # Se for um executável .exe
        base = os.path.dirname(sys.executable)  # Diretório onde o .exe está rodando
    else:
        base = os.path.dirname(os.path.abspath(__file__))  # Caminho normal no Python

    return str(os.path.join(base, caminho)) # Retorna o caminho correto para templates/ como uma string comum

def criar_pasta_saida(num_portaria_completo: str) -> str:
    # Usando expressão regular para capturar os números iniciais do num_portaria_completo
    resultado = re.match(r'(\d+)/IPM/CORREG/PMMS/', num_portaria_completo)
    num = 0
    if resultado:
        num = resultado.group(1)  # Retorna apenas os números capturados

    def get_caminho_base():
        """ Retorna o caminho correto da pasta base, dentro ou fora do PyInstaller """
        if getattr(sys, 'frozen', False):  # Se for um executável .exe
            return os.path.dirname(sys.executable)  # Diretório onde o .exe está
        return os.path.dirname(os.path.abspath(__file__))  # Diretório do script Python

    # Obtém o caminho base correto
    caminho_principal = get_caminho_base()

    # Nome da nova pasta, num_portaria_completo[-4:] pega o ano 2024 que fica no final de /IPM/CORREG/PMMS/2024
    nome_pasta = "IPM Portaria n " + str(num) + " " + num_portaria_completo[-4:]

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
    # Caminho correto para o template
    caminho_template = get_caminho_templates("templates/ipm/A - Capa.docx")

    # Carregar o template
    documento = DocxTemplate(caminho_template)
    dados = {
        "uopm": obj["uopm"],
        "num_portaria": obj["num_portaria"],
        "data_portaria": obj["data_portaria"],
        "posto_encarregado": obj["posto_encarregado"],
        "nome_encarregado": obj["nome_encarregado"],
        "mat_encarregado": obj["mat_encarregado"],
        "postograd_investigado": obj["postograd_investigado"],
        "nome_investigado": obj["nome_investigado"],
        "mat_investigado": obj["mat_investigado"],
        "texto_finalidade": obj["texto_finalidade"]
    }
    documento.render(dados)
    nome_da_pasta = criar_pasta_saida(obj["num_portaria"])
    caminho_saida = nome_da_pasta + r"\A - Capa.docx"
    documento.save(caminho_saida)
    return None


def processar_autuacao(obj: Dict[str, str]) -> None:
    # Caminho correto para o template
    caminho_template = get_caminho_templates("templates/ipm/B - Autuacao.docx")

    # Carregar o template
    documento = DocxTemplate(caminho_template)

    dados = {
        "num_portaria": obj["num_portaria"],
        "data_portaria": obj["data_portaria"],
        "uopm": obj["uopm"],
        "grande_comando": obj["grande_comando"],
        "uopm_extenso": obj["uopm_extenso"],
        "uopm_email": obj["uopm_email"],
        "uopm_telefone": obj["uopm_telefone"],
        "uopm_cidade": obj["uopm_cidade"],
        "uopm_endereco": obj["uopm_endereco"],  # dados do cabeçalho e rodapé terminam aqui
        "posto_encarregado": obj["posto_encarregado"],
        "nome_encarregado": obj["nome_encarregado"],
        "mat_encarregado": obj["mat_encarregado"],
        "postograd_escrivao": obj["postograd_escrivao"],
        "nome_escrivao": obj["nome_escrivao"],
        "mat_escrivao": obj["mat_escrivao"],
        "postograd_investigado": obj["postograd_investigado"],
        "nome_investigado": obj["nome_investigado"],
        "mat_investigado": obj["mat_investigado"],
        "data_autuacaoextenso": obj["data_autuacaoextenso"],
    }
    documento.render(dados)
    nome_da_pasta = criar_pasta_saida(obj["num_portaria"])
    caminho_saida = nome_da_pasta + r"\B - Autuação.docx"
    documento.save(caminho_saida)
    return None


def processar_portaria_encarregado(obj: Dict[str, str]) -> None:
    # Caminho correto para o template
    caminho_template = get_caminho_templates("templates/ipm/C - Portaria do encarregado.docx")

    # Carregar o template
    documento = DocxTemplate(caminho_template)

    dados = {
        "uopm_cidade": obj["uopm_cidade"],
        "uopm_extenso": obj["uopm_extenso"],
        "num_portaria": obj["num_portaria"],
        "data_portaria": obj["data_portaria"],
        "uopm": obj["uopm"],
        "grande_comando": obj["grande_comando"],
        "uopm_email": obj["uopm_email"],
        "uopm_telefone": obj["uopm_telefone"],
        "uopm_endereco": obj["uopm_endereco"],  # dados do cabeçalho e rodapé terminam aqui
        "posto_autinst": obj["posto_autinst"],
        "nome_autinst": obj["nome_autinst"],
        "func_autinst": obj["func_autinst"],
        "texto_finalidade": obj["texto_finalidade"],
        "data_autuacao": obj["data_autuacao"],
        "nome_encarregado": obj["nome_encarregado"],
        "posto_encarregado": obj["posto_encarregado"],
        "mat_encarregado": obj["mat_encarregado"]
    }

    documento.render(dados)
    nome_da_pasta = criar_pasta_saida(obj["num_portaria"])
    caminho_saida = nome_da_pasta + r"\C - Portaria do encarregado.docx"
    documento.save(caminho_saida)
    return None


def processar_designacao(obj: Dict[str, str]) -> None:
    # Caminho correto para o template
    caminho_template = get_caminho_templates("templates/ipm/D - Designacao escrivao.docx")

    # Carregar o template
    documento = DocxTemplate(caminho_template)

    dados = {
        "uopm_cidade": obj["uopm_cidade"],
        "uopm_extenso": obj["uopm_extenso"],
        "num_portaria": obj["num_portaria"],
        "data_portaria": obj["data_portaria"],
        "uopm": obj["uopm"],
        "grande_comando": obj["grande_comando"],
        "uopm_email": obj["uopm_email"],
        "uopm_telefone": obj["uopm_telefone"],
        "uopm_endereco": obj["uopm_endereco"],  # dados do cabeçalho e rodapé terminam aqui
        "postograd_escrivao": obj["postograd_escrivao"],
        "nome_escrivao": obj["nome_escrivao"],
        "mat_escrivao": obj["mat_escrivao"],
        "data_autuacao": obj["data_autuacao"],
        "nome_encarregado": obj["nome_encarregado"],
        "posto_encarregado": obj["posto_encarregado"],
        "mat_encarregado": obj["mat_encarregado"]
    }
    documento.render(dados)
    nome_da_pasta = criar_pasta_saida(obj["num_portaria"])
    caminho_saida = nome_da_pasta + r"\D - Designação escrivao.docx"
    documento.save(caminho_saida)
    return None


def processar_termo_compromisso(obj: Dict[str, str]) -> None:
    # Caminho correto para o template
    caminho_template = get_caminho_templates("templates/ipm/E - Termo de compromisso.docx")

    # Carregar o template
    documento = DocxTemplate(caminho_template)
    dados = {
        "uopm_cidade": obj["uopm_cidade"],
        "uopm_extenso": obj["uopm_extenso"],
        "num_portaria": obj["num_portaria"],
        "data_portaria": obj["data_portaria"],
        "uopm": obj["uopm"],
        "grande_comando": obj["grande_comando"],
        "uopm_email": obj["uopm_email"],
        "uopm_telefone": obj["uopm_telefone"],
        "uopm_endereco": obj["uopm_endereco"],  # dados do cabeçalho e rodapé terminam aqui
        "posto_encarregado": obj["posto_encarregado"],
        "nome_encarregado": obj["nome_encarregado"],
        "mat_encarregado": obj["mat_encarregado"],
        "postograd_escrivao": obj["postograd_escrivao"],
        "nome_escrivao": obj["nome_escrivao"],
        "mat_escrivao": obj["mat_escrivao"],
        "postograd_investigado": obj["postograd_investigado"],
        "nome_investigado": obj["nome_investigado"],
        "mat_investigado": obj["mat_investigado"],
        "data_autuacaoextenso": obj["data_autuacaoextenso"],
    }
    documento.render(dados)
    nome_da_pasta = criar_pasta_saida(obj["num_portaria"])
    caminho_saida = nome_da_pasta + r"\E - Termo de compromisso.docx"
    documento.save(caminho_saida)
    return None


def processar_despacho(obj: Dict[str, str]) -> None:
    # Caminho correto para o template
    caminho_template = get_caminho_templates("templates/ipm/F - Despacho 001.docx")

    # Carregar o template
    documento = DocxTemplate(caminho_template)
    dados = {
        "num_portaria": obj["num_portaria"],
        "data_portaria": obj["data_portaria"],
        "uopm": obj["uopm"],
        "grande_comando": obj["grande_comando"],
        "uopm_extenso": obj["uopm_extenso"],
        "uopm_email": obj["uopm_email"],
        "uopm_telefone": obj["uopm_telefone"],
        "uopm_cidade": obj["uopm_cidade"],
        "uopm_endereco": obj["uopm_endereco"],  # dados do cabeçalho e rodapé terminam aqui
        "func_autinst": obj["func_autinst"],
        "posto_encarregado": obj["posto_encarregado"],
        "nome_encarregado": obj["nome_encarregado"],
        "mat_encarregado": obj["mat_encarregado"],
        "data_autuacao": obj["data_autuacao"],
    }
    documento.render(dados)
    nome_da_pasta = criar_pasta_saida(obj["num_portaria"])
    caminho_saida = nome_da_pasta + r"\F - Despacho 001.docx"
    documento.save(caminho_saida)
    return None


def processar_recebimento(obj: Dict[str, str]) -> None:
    # Caminho correto para o template
    caminho_template = get_caminho_templates("templates/ipm/G - Recebimento.docx")

    # Carregar o template
    documento = DocxTemplate(caminho_template)
    dados = {
        "grande_comando": obj["grande_comando"],
        "uopm_extenso": obj["uopm_extenso"],
        "uopm_email": obj["uopm_email"],
        "uopm_telefone": obj["uopm_telefone"],
        "uopm_cidade": obj["uopm_cidade"],
        "uopm_endereco": obj["uopm_endereco"],  # dados do cabeçalho e rodapé terminam aqui
        "data_autuacaoextenso": obj["data_autuacaoextenso"],
        "uopm": obj["uopm"],
        "postograd_escrivao": obj["postograd_escrivao"],
        "nome_escrivao": obj["nome_escrivao"],
        "mat_escrivao": obj["mat_escrivao"],
    }
    documento.render(dados)
    nome_da_pasta = criar_pasta_saida(obj["num_portaria"])
    caminho_saida = nome_da_pasta + r"\G - Recebimento.docx"
    documento.save(caminho_saida)
    return None


def processar_oficio_001(obj: Dict[str, str]) -> None:
    # Caminho correto para o template
    caminho_template = get_caminho_templates("templates/ipm/G.1 - Of 001 informando designacao.docx")

    # Carregar o template
    documento = DocxTemplate(caminho_template)
    dados = {
        "uopm": obj["uopm"],
        "grande_comando": obj["grande_comando"],
        "uopm_extenso": obj["uopm_extenso"],
        "uopm_email": obj["uopm_email"],
        "uopm_telefone": obj["uopm_telefone"],
        "uopm_cidade": obj["uopm_cidade"],
        "uopm_endereco": obj["uopm_endereco"],  # dados do cabeçalho e rodapé terminam aqui
        "num_portaria": obj["num_portaria"],
        "data_autuacao": obj["data_autuacao"],
        "func_autinst": obj["func_autinst"],
        "data_portaria": obj["data_portaria"],
        "postograd_escrivao": obj["postograd_escrivao"],
        "nome_escrivao": obj["nome_escrivao"],
        "mat_escrivao": obj["mat_escrivao"],
        "posto_encarregado": obj["posto_encarregado"],
        "nome_encarregado": obj["nome_encarregado"],
        "mat_encarregado": obj["mat_encarregado"],
    }
    documento.render(dados)
    nome_da_pasta = criar_pasta_saida(obj["num_portaria"])
    caminho_saida = nome_da_pasta + r"\G.1 - Of 001 informando designacao.docx"
    documento.save(caminho_saida)
    return None


def processar_oitiva_investigado(obj: Dict[str, str]) -> None:
    # Caminho correto para o template
    caminho_template = get_caminho_templates("templates/ipm/G.2 - Oitiva investigado.docx")

    # Carregar o template
    documento = DocxTemplate(caminho_template)

    dados = {
        "uopm": obj["uopm"],
        "grande_comando": obj["grande_comando"],
        "uopm_extenso": obj["uopm_extenso"],
        "uopm_email": obj["uopm_email"],
        "uopm_telefone": obj["uopm_telefone"],
        "uopm_cidade": obj["uopm_cidade"],
        "uopm_endereco": obj["uopm_endereco"],  # dados do cabeçalho e rodapé terminam aqui
        "data_autuacao": obj["data_autuacao"],
        "posto_encarregado": obj["posto_encarregado"],
        "nome_encarregado": obj["nome_encarregado"],
        "mat_encarregado": obj["mat_encarregado"],
        "postograd_escrivao": obj["postograd_escrivao"],
        "nome_escrivao": obj["nome_escrivao"],
        "mat_escrivao": obj["mat_escrivao"],
        "num_portaria": obj["num_portaria"],
        "data_portaria": obj["data_portaria"],
        "nome_investigado": obj["nome_investigado"],
        "mat_investigado": obj["mat_investigado"],
        "postograd_investigado": obj["postograd_investigado"]
    }
    documento.render(dados)
    nome_da_pasta = criar_pasta_saida(obj["num_portaria"])
    caminho_saida = nome_da_pasta + r"\G.2 - Oitiva investigado.docx"
    documento.save(caminho_saida)
    return None


def processar_conclusao(obj: Dict[str, str]) -> None:
    # Caminho correto para o template
    caminho_template = get_caminho_templates("templates/ipm/V - Conclusao.docx")

    # Carregar o template
    documento = DocxTemplate(caminho_template)
    dados = {
        "uopm": obj["uopm"],
        "grande_comando": obj["grande_comando"],
        "uopm_extenso": obj["uopm_extenso"],
        "uopm_email": obj["uopm_email"],
        "uopm_telefone": obj["uopm_telefone"],
        "uopm_cidade": obj["uopm_cidade"],
        "uopm_endereco": obj["uopm_endereco"],  # dados do cabeçalho e rodapé terminam aqui
        "data_relatorio": obj["data_relatorio"],
        "postograd_escrivao": obj["postograd_escrivao"],
        "nome_escrivao": obj["nome_escrivao"],
        "mat_escrivao": obj["mat_escrivao"],
    }
    documento.render(dados)
    nome_da_pasta = criar_pasta_saida(obj["num_portaria"])
    caminho_saida = nome_da_pasta + r"\V - Conclusao.docx"
    documento.save(caminho_saida)
    return None


def processar_relatorio(obj: Dict[str, str]) -> None:
    # Caminho correto para o template
    caminho_template = get_caminho_templates("templates/ipm/W - Relatorio.docx")

    # Carregar o template
    documento = DocxTemplate(caminho_template)
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
        "postograd_investigado": obj["postograd_investigado"],
        "nome_investigado": obj["nome_investigado"],
        "mat_investigado": obj["mat_investigado"],
        "data_relatorio": obj["data_relatorio"],
        "nome_encarregado": obj["nome_encarregado"],
        "posto_encarregado": obj["posto_encarregado"],
        "mat_encarregado": obj["mat_encarregado"]
    }
    documento.render(dados)
    nome_da_pasta = criar_pasta_saida(obj["num_portaria"])
    caminho_saida = nome_da_pasta + r"\W - Relatorio.docx"
    documento.save(caminho_saida)
    return None


def processar_termo_remessa(obj: Dict[str, str]) -> None:
    # Caminho correto para o template
    caminho_template = get_caminho_templates("templates/ipm/X - Termo de remessa.docx")

    # Carregar o template
    documento = DocxTemplate(caminho_template)
    dados = {
        "uopm_cidade": obj["uopm_cidade"],
        "uopm_extenso": obj["uopm_extenso"],
        "num_portaria": obj["num_portaria"],
        "data_portaria": obj["data_portaria"],
        "uopm": obj["uopm"],
        "grande_comando": obj["grande_comando"],
        "uopm_email": obj["uopm_email"],
        "uopm_telefone": obj["uopm_telefone"],
        "uopm_endereco": obj["uopm_endereco"],  # dados do cabeçalho e rodapé terminam aqui
        "posto_encarregado": obj["posto_encarregado"],
        "nome_encarregado": obj["nome_encarregado"],
        "mat_encarregado": obj["mat_encarregado"],
        "data_relatorio": obj["data_relatorio"],
        "func_autinst": obj["func_autinst"]
    }
    documento.render(dados)
    nome_da_pasta = criar_pasta_saida(obj["num_portaria"])
    caminho_saida = nome_da_pasta + r"\X - Termo de remessa.docx"
    documento.save(caminho_saida)
    return None


def processar_oficio_remessa(obj: Dict[str, str]) -> None:
    # Caminho correto para o template
    caminho_template = get_caminho_templates("templates/ipm/Y - Of de remessa.docx")

    # Carregar o template
    documento = DocxTemplate(caminho_template)
    dados = {
        "uopm": obj["uopm"],
        "grande_comando": obj["grande_comando"],
        "uopm_extenso": obj["uopm_extenso"],
        "uopm_email": obj["uopm_email"],
        "uopm_telefone": obj["uopm_telefone"],
        "uopm_cidade": obj["uopm_cidade"],
        "uopm_endereco": obj["uopm_endereco"], #dados do cabeçalho e rodapé até aqui
        "num_portaria": obj["num_portaria"],
        "data_relatorio": obj["data_relatorio"],
        "func_autinst": obj["func_autinst"],
        "data_portaria": obj["data_portaria"],
        "nome_encarregado": obj["nome_encarregado"],
        "posto_encarregado": obj["posto_encarregado"],
        "mat_encarregado": obj["mat_encarregado"]
    }
    documento.render(dados)
    nome_da_pasta = criar_pasta_saida(obj["num_portaria"])
    caminho_saida = nome_da_pasta + r"\Y - Of de remessa.docx"
    documento.save(caminho_saida)
    return None

def processar_solucao(obj: Dict[str, str]) -> None:
    # Caminho correto para o template
    caminho_template = get_caminho_templates("templates/ipm/Z - Solucao para o comandante.docx")

    # Carregar o template
    documento = DocxTemplate(caminho_template)
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
        "postograd_investigado": obj["postograd_investigado"],
        "nome_investigado": obj["nome_investigado"],
        "mat_investigado": obj["mat_investigado"],
        "data_relatorio": obj["data_relatorio"],
        "nome_encarregado": obj["nome_encarregado"],
        "posto_encarregado": obj["posto_encarregado"],
        "mat_encarregado": obj["mat_encarregado"]
    }
    documento.render(dados)
    nome_da_pasta = criar_pasta_saida(obj["num_portaria"])
    caminho_saida = nome_da_pasta + r"\Z - Solucao para o comandante.docx"
    documento.save(caminho_saida)
    return None
