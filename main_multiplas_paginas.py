import re
import pypdf
import pandas as pd
from datetime import datetime
from os import path


pdf_path = "notas.pdf"
arquivo_destino = 'notas'

# 1. Expressão Regular para capturar as linhas da tabela
#
# Explicação do Regex:
# (BOVESPA) -> Procura pela palavra Mercado (ex: BOVESPA)
# \s+(C|V)\s+(VISTA|FRACIONARIO|TERMO) -> Captura C ou V e o tipo de mercado
# \s+([\w\s\d]+(?:\s+CI|\s+ON|\s+PN)?) -> Captura o nome do ativo (ex: BRCR11 CI, BOVA11 CI)
# .*? -> Ignora a coluna de Observação vazia
# \s+(\d+)\s+ -> Captura a quantidade (número)
# R\$\s*([\d.,]+)\s+ -> Captura o preço unitário
# R\$\s*([\d.,]+)\s+(D|C) -> Captura o valor total e se é Débito ou Crédito
#
padrao_linha = re.compile(
    r"(BOVESPA)\s+(C|V)\s+(VISTA|FRACIONARIO|TERMO)\s+(.*?)\s+(\d+)\s+R\$\s*([\d.,]+)\s+R\$\s*([\d.,]+)\s+(D|C)"
)

padrao_data = re.compile(r"Data Pregão\s+(\d{2}/\d{2}/\d{4})")

# Lista para acumular os dados de TODAS as páginas
todos_os_dados = []

print(f"Abrindo o arquivo '{pdf_path}'...")

# 2. Ler o PDF e percorrer todas as páginas
with open(pdf_path, "rb") as f:
    leitor = pypdf.PdfReader(f)
    total_paginas = len(leitor.pages)
    print(f"Total de páginas encontradas: {total_paginas}")
    
    for i, pagina in enumerate(leitor.pages):
        print(f"Processando página {i + 1}/{total_paginas}...")
        
        # Extrai o texto da página atual
        texto_pagina = pagina.extract_text()
        
        # Procura a data do pregão nesta página. 
        # Se não achar na página atual, tenta manter a da página anterior (caso seja uma continuação)
        busca_data = padrao_data.search(texto_pagina)
        data_pregao = busca_data.group(1) if busca_data else "Não encontrada"
        
        # Procura as linhas de operações
        linhas_encontradas = padrao_linha.findall(texto_pagina)

        # Se encontrou linhas nesta página, processa e adiciona à lista geral
        for item in linhas_encontradas:
            mercado = item[0]
            cv = f"{item[1]}"
            tipo = f"{item[2]}"
            especificacao = item[3].strip()[0:6]
            observacao = item[3][6:].strip()        # Pegar dados da coluna anterior sem o ticker
            quantidade = item[4]
            preco = f"{item[5]}"
            valor = f"{item[6]}"
            dc = f"{item[7]}"
            
            todos_os_dados.append(
                [
                    data_pregao,
                    mercado,
                    cv,
                    tipo,
                    especificacao,
                    observacao,
                    quantidade,
                    preco,
                    valor,
                    dc
                ]
            )

# 3. Se acumulou dados de qualquer uma das páginas, estrutura no Pandas
if todos_os_dados:
    # Nomes exatos das colunas desejadas
    colunas = [
        "Data Pregao",
        "Mercado", 
        "C/V",
        "Tipo de Mercado",
        "Especificação do Título", 
        "Observação",
        "Quantidade",
        "Preço/Ajuste (R$)",
        "Valor/Ajuste (R$)",
        "D/C"
    ]
    
    df_resultado = pd.DataFrame(todos_os_dados, columns=colunas)
    
    print("\n--- Processamento Concluído com Sucesso ---")
    print(f"Total de operações extraídas: {len(df_resultado)}")
    print("\n--- Visualização dos Dados ---")
    print(df_resultado.to_string(index=False))
    
    # Gravar todas as páginas consolidadas em um único arquivo Excel
    data_hora_atual = datetime.strftime(datetime.now(), '%d%m%Y_%H%M%S')
    arquivo_destino = path.splitext(arquivo_destino)[0] + '_' + data_hora_atual
    df_resultado.to_excel(arquivo_destino + '.xlsx', index=False)
    print(f"\nArquivo [ {arquivo_destino} ] gerado com sucesso!")

else:
    print("\nNenhuma linha correspondente ao padrão da tabela foi encontrada em nenhuma das páginas.")