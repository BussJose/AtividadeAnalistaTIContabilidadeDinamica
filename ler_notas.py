from pathlib import Path  # ferramenta para trabalhar com pastas e arquivos
import xml.etree.ElementTree as ET  # ferramenta para ler arquivos XML
import pandas as pd  # ferramenta para criar tabelas e salvar em Excel

PASTA = Path("notas")  # pasta onde ficam os XMLs
NS = {"nfe": "http://www.portalfiscal.inf.br/nfe"}  # endereço (namespace) que toda NF-e usa


def texto(no, caminho):  # função que busca um campo dentro da nota
    """Busca um campo dentro da nota; devolve '' se não existir."""
    achado = no.find(caminho, NS)  # procura o campo no XML
    return achado.text if achado is not None else ""  # devolve o texto, ou vazio se não achou

itens = []
linhas = []  # lista vazia: cada nota lida vira uma linha aqui
for arquivo in PASTA.glob("*.xml"):  # repete para cada arquivo .xml da pasta
    try:  # tenta ler a nota; se der erro, vai para o except
        raiz = ET.parse(arquivo).getroot()  # abre o XML e pega o elemento principal
        inf = raiz.find(".//nfe:infNFe", NS)  # acha o bloco com os dados da nota
        linhas.append({  # adiciona uma nova linha na lista
            "arquivo": arquivo.name,  # nome do arquivo XML
            "numero": texto(inf, "nfe:ide/nfe:nNF"),  # número da nota
            "dt_emissao": texto(inf, "nfe:ide/nfe:dhEmi")[:10],  # data de emissão (10 primeiros caracteres, sem a hora)
            "emitente": texto(inf, "nfe:emit/nfe:xNome") or "NÃO INFORMADO",  # nome de quem emitiu a nota
            "cnpj_emitente": texto(inf, "nfe:emit/nfe:CNPJ"),  # CNPJ de quem emitiu
            "valor_produtos": float(texto(inf, "nfe:total/nfe:ICMSTot/nfe:vProd") or 0),  # valor dos produtos (número; 0 se vazio)
            "icms": float(texto(inf, "nfe:total/nfe:ICMSTot/nfe:vICMS") or 0),  # valor do ICMS
            "valor_total": float(texto(inf, "nfe:total/nfe:ICMSTot/nfe:vNF") or 0),  # valor total da nota
        })
        for item in inf.findall("nfe:det", NS):
            itens.append({
                "numero_nota": texto(inf, "nfe:ide/nfe:nNF"),
                "produto": texto(item, "nfe:prod/nfe:xProd"),
                "ncm": texto(item, "nfe:prod/nfe:NCM"),
                "quantidade": float(texto(item, "nfe:prod/nfe:qCom") or 0),
                "valor_item": float(texto(item, "nfe:prod/nfe:vProd") or 0),
            })        
    except Exception as erro:  # se o arquivo estiver quebrado, cai aqui
        print(f"Erro em {arquivo.name}: {erro}")  # mostra qual arquivo deu erro e segue em frente

df = pd.DataFrame(linhas)  # transforma a lista em uma tabela
df_itens = pd.DataFrame(itens)
resumo = df.groupby("emitente", as_index=False)[["valor_total", "icms"]].sum()  # soma valor e ICMS por emitente

with pd.ExcelWriter("resumo_notas.xlsx") as saida:  # cria o arquivo Excel (feche-o no Excel antes de rodar)
    df.to_excel(saida, sheet_name="Notas", index=False)  # grava a tabela completa na aba "Notas"
    df_itens.to_excel(saida, sheet_name="Itens", index=False)
    resumo.to_excel(saida, sheet_name="Resumo por emitente", index=False)  # grava o resumo em outra aba

print(f"{len(df)} notas processadas. Arquivo resumo_notas.xlsx criado.")  # avisa quantas notas foram lidas
