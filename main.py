from zeep import Client
from pathlib import Path
import os
import configparser

from zeep.xsd.types.builtins import String
# 1. Criar um script python para fazer a leitura de estrutura de pastas que existe no fluig e criar no windows;
# 2. Ajustar o script para baixar os arquivos dentro das pastas;
# 3. Ajustar o script para baixar os formulários no formato de pasta e os registros de formulário em CSV;
# 4. Ajustar o script para criar uma planilha catalogando o caminho dos arquivos e as informações do ECM Fluig;
# 5. Ajustar o script para ser um CLI

def obter_documento(info_ambiente, pasta_id, tipo_documento):
    SERVICO = Client(info_ambiente["host"] + "webdesk/ECMDatasetService?wsdl")

    constraint_Array = SERVICO.get_type('ns0:searchConstraintDtoArray')
    constraint = SERVICO.get_type('ns0:searchConstraintDto')
    tipos = {
        "PastaRaiz": "0",
        "Pasta": "1",
        "DocumentoNormal": "2",
        "DocumentoExterno": "3",
        "Formulario": "4",
        "RegistroFormulario": "5",
        "AnexoWorkflow": "7",
        "NovoConteudo": "8",
        "Aplicativo": "9",
        "Relatorio": "10"
    }

    constraints = []
    constraints.append(constraint("MUST", "documentType", tipos[tipo_documento], tipos[tipo_documento], False))
    constraints.append(constraint("MUST", "parentDocumentId", pasta_id, pasta_id, False))
    constraints.append(constraint("MUST", "deleted", "false", "false", False))

    return SERVICO.service.getDataset(
        info_ambiente["empresa_id"],
        info_ambiente["ususario"],
        info_ambiente["senha"],
        "document",
        "",
        constraint_Array(constraints),
        ""
    )

def obter_anexos(info_ambiente, documento_id, versao):
    SERVICO_CARDINDEX = Client(info_ambiente["host"] + "webdesk/ECMCardIndexService?wsdl")
    SERVICO_DOCUMENT = Client(info_ambiente["host"] + "webdesk/ECMDocumentService?wsdl")

    nome_dos_anexos = SERVICO_CARDINDEX.service.getAttachmentsList(
        info_ambiente["ususario"],
        info_ambiente["senha"],
        info_ambiente["empresa_id"],
        documento_id
    )

    anexos = []
    for nome in nome_dos_anexos:
        anexo = SERVICO_DOCUMENT.service.getDocumentContent(
            info_ambiente["ususario"],
            info_ambiente["senha"],
            info_ambiente["empresa_id"],
            documento_id,
            info_ambiente["matricula"],
            versao,
            nome
        )
        anexos.append({
            "nome": nome,
            "base64": anexo
        })
    return anexos

def criar_pasta(info_ambiente, parente_id, caminho):
    pastas = obter_documento(info_ambiente, parente_id, "Pasta")
    for pasta in pastas.values:
        pasta_id = pasta.value[pastas.columns.index("documentPK.documentId")]
        pasta_nome = pasta.value[pastas.columns.index("documentDescription")]

        if not pasta_nome:
            pasta_nome = "Nova pasta"

        sub_caminho = caminho / "{} - {}".format(pasta_id, pasta_nome)
        sub_caminho.mkdir(parents=True, exist_ok=True)

        print("========= CRIAÇÃO DA PASTA =========\nPARENTE ID: {}\nID: {}\nCAMINHO: {}\n\n".format(parente_id, pasta_id, sub_caminho))
        criar_pasta(info_ambiente, pasta_id, sub_caminho)
        criar_documento_normal(info_ambiente, pasta_id, sub_caminho)

def criar_documento_normal(info_ambiente, parente_id, caminho):
    documentos = obter_documento(info_ambiente, parente_id, "DocumentoNormal")
    for documento in documentos.values:
        codigo = documento.value[documentos.columns.index("documentPK.documentId")]
        versao = documento.value[documentos.columns.index("documentPK.version")]
        nome = documento.value[documentos.columns.index("documentDescription")]

        if not nome:
            nome = "Novo documento"

        sub_caminho = caminho / "{} - {}".format(codigo, nome)
        sub_caminho.mkdir(parents=True, exist_ok=True)

        print("======= CRIAÇÃO DO DOCUMENTO =======\nPARENTE ID: {}\nID: {}\nCAMINHO: {}\n\n".format(parente_id, codigo, sub_caminho))
        anexos = obter_anexos(info_ambiente, codigo, versao)

        print("========= CRIAÇÃO DO ANEXO =========")
        for anexo in anexos:
            anexo_nome = sub_caminho / anexo["nome"]
            print("ANEXO: {}".format(anexo_nome))
            arquivo = open(anexo_nome, "wb")
            arquivo.write(anexo["base64"])
            arquivo.close()
        print("\n\n")

def criar_documento_externo(info_ambiente, parente_id, caminho):
    documentos = obter_documento(info_ambiente, parente_id, "DocumentoExterno")
    for documento in documentos.values:
        codigo = documento.value[documentos.columns.index("documentPK.documentId")]
        nome = documento.value[documentos.columns.index("documentDescription")]
        url = documento.value[documentos.columns.index("phisicalFile")]
        descricao = documento.value[documentos.columns.index("additionalComments")]

        if not nome:
            nome = "Novo documento externo"

        sub_caminho = caminho / "{} - {}".format(codigo, nome)

        print("======= CRIAÇÃO DO DOCUMENTO =======\nPARENTE ID: {}\nID: {}\nCAMINHO: {}\n\n".format(parente_id, codigo, sub_caminho))
        texto = f"""
        URL: {url}
        Descrição: {descricao}
        """

        arquivo = open(sub_caminho, "wb")
        arquivo.write(texto)
        arquivo.close()

def criar_formulario(info_ambiente, parente_id, caminho):
    documentos = obter_documento(info_ambiente, parente_id, "Formulario")
    for documento in documentos.values:
        codigo = documento.value[documentos.columns.index("documentPK.documentId")]
        nome = documento.value[documentos.columns.index("documentDescription")]
        dataset = documento.value[documentos.columns.index("datasetName")]

        if not nome:
            nome = "Novo formulário"

        sub_caminho = caminho / "{} - {}".format(codigo, nome)

        print("======= CRIAÇÃO DO FORMULÁRIO =======\nPARENTE ID: {}\nID: {}\nCAMINHO: {}\n\n".format(parente_id, codigo, sub_caminho))


        arquivo = open(sub_caminho, "wb")
        arquivo.write(texto)
        arquivo.close()

def main():
    cfg = configparser.ConfigParser()
    cfg.read("config.ini")

    info_ambiente = {
        "host": cfg.get("ambiente", "host"),
        "empresa_id": cfg.getint("ambiente", "empresaid"),
        "matricula": cfg.get("ambiente", "matricula"),
        "ususario": cfg.get("ambiente", "ususario"),
        "senha": cfg.get("ambiente", "senha")
    }

    caminho = Path("Documentos")
    caminho.mkdir(parents=True, exist_ok=True)
    criar_pasta(info_ambiente, "0", caminho)

if __name__ == "__main__":
    main()