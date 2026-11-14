# Leitor de XML de NF-e para Excel

Script em Python que lê notas fiscais eletrônicas (NF-e) em XML e gera uma planilha Excel com os dados organizados, o resumo por emitente e os itens de cada nota.

Projeto de automação criado para reduzir o trabalho manual de digitar ou conferir notas fiscais na rotina de um escritório de contabilidade.

## O que o script faz

- Lê todos os arquivos `.xml` de uma pasta
- Extrai de cada nota: número, data de emissão, emitente, CNPJ, valor dos produtos, ICMS e valor total
- Extrai os itens de cada nota: produto, NCM, quantidade e valor
- Soma valor total e ICMS por emitente
- Gera o arquivo `resumo_notas.xlsx` com três abas
- Trata erros: arquivo XML corrompido não interrompe o processamento das demais notas, e campos ausentes (como emitente) não travam o script

## Planilha gerada

| Aba | Conteúdo |
|---|---|
| **Notas** | Uma linha por nota fiscal |
| **Itens** | Uma linha por produto de cada nota |
| **Resumo por emitente** | Total de valor e ICMS agrupado por emitente |

## Tecnologias

- Python 3
- pandas
- openpyxl
- xml.etree.ElementTree (biblioteca padrão)

## Como usar

1. Instale o [Python](https://www.python.org/downloads/) (marque a opção **Add Python to PATH** na instalação).
2. Instale as dependências:
   ```
   pip install pandas openpyxl
   ```
3. Coloque os arquivos XML na pasta `notas`.
4. Execute:
   ```
   python ler_notas.py
   ```
5. Abra o arquivo `resumo_notas.xlsx` gerado na mesma pasta.

> Feche o `resumo_notas.xlsx` no Excel antes de executar o script novamente, senão o Windows bloqueia a gravação.

## Estrutura do projeto

```
.
├── ler_notas.py       # script principal
├── notas/             # XMLs de exemplo
│   ├── nfe_01.xml ... nfe_08.xml
│   ├── nfe_09_quebrada.xml        # arquivo corrompido, para testar tratamento de erro
│   └── nfe_10_sem_emitente.xml    # nota sem emitente, para testar campo vazio
└── README.md
```

## Casos de teste incluídos

- **Notas válidas** (`nfe_01` a `nfe_08`): processadas normalmente.
- **XML corrompido** (`nfe_09_quebrada`): o script exibe o erro e continua com as demais notas.
- **Nota sem emitente** (`nfe_10_sem_emitente`): o emitente aparece como "NÃO INFORMADO".

## Aviso

Todas as notas da pasta `notas` são **fictícias**, criadas apenas para demonstração. Nunca publique XMLs reais de clientes em repositórios públicos.

## Próximos passos

- Dashboard no Looker Studio ou Power BI a partir da planilha gerada
- Agente de IA para responder perguntas sobre as notas em linguagem natural
- Suporte a NFS-e (que varia por prefeitura)
- Envio automático do resumo por e-mail

## Autor

José Augusto Buss
[LinkedIn](https://www.linkedin.com/in/jos%C3%A9-buss-084655223/)
