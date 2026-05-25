import re
import pypdf
import pandas as pd

pdf_path = "notas.pdf"

# 1. Extrair o texto bruto do PDF
print("Extraindo o texto do PDF...")
with open(pdf_path, "rb") as f:
    leitor = pypdf.PdfReader(f)
    # Pega o texto da primeira página
    texto_bruto = leitor.pages[0].extract_text()

# 2. Expressão Regular para capturar as linhas da tabela
# Explicação do Regex:
# (BOVESPA) -> Procura pela palavra Mercado (ex: BOVESPA)
# \s+(C|V)\s+(VISTA|FRACIONARIO|TERMO) -> Captura C ou V e o tipo de mercado
# \s+([\w\s\d]+(?:\s+CI|\s+ON|\s+PN)?) -> Captura o nome do ativo (ex: BRCR11 CI, BOVA11 CI)
# .*? -> Ignora a coluna de Observação vazia
# \s+(\d+)\s+ -> Captura a quantidade (número)
# R\$\s*([\d.,]+)\s+ -> Captura o preço unitário
# R\$\s*([\d.,]+)\s+(D|C) -> Captura o valor total e se é Débito ou Crédito
padrao_linha = re.compile(
    r"(BOVESPA)\s+(C|V)\s+(VISTA|FRACIONARIO|TERMO)\s+(.*?)\s+(\d+)\s+R\$\s*([\d.,]+)\s+R\$\s*([\d.,]+)\s+(D|C)"
)

# Encontrar todas as ocorrências no texto
linhas_encontradas = padrao_linha.findall(texto_bruto)

# 3. Se encontrou dados, estruturar no Pandas
if linhas_encontradas:
    dados_processados = []
    
    for item in linhas_encontradas:
        # Organizando as colunas capturadas pelo Regex de forma idêntica ao PDF
        mercado = item[0]
        cv_tipo = f"{item[1]} {item[2]}" # Junta o 'C' ou 'V' com 'VISTA' -> 'C VISTA'
        especificacao = item[3].strip()
        observacao = "" # Fica em branco conforme o padrão da nota
        quantidade = item[4]
        preco = f"R$ {item[5]}"
        valor_ajuste = f"R$ {item[6]} {item[7]}"
        
        dados_processados.append([
            mercado, cv_tipo, especificacao, observacao, quantidade, preco, valor_ajuste
        ])
    
    # Criar o DataFrame com os nomes exatos das colunas que você deseja
    colunas = [
        "Mercado", "C/V Tipo de Mercado", "Especificação do Título", 
        "Observação", "Quantidade", "Preço/Ajuste", "Valor/Ajuste D/C"
    ]
    
    df_resultado = pd.DataFrame(dados_processados, columns=colunas)
    
    print("\n--- Tabela Extraída via REGEX com Sucesso ---")
    print(df_resultado.to_string(index=False))
    
    # Gravar em arquivo Excel
    df_resultado.to_excel("tabela_nota_regex.xlsx", index=False)
    print("\nArquivo 'tabela_nota_regex.xlsx' gerado!")

else:
    print("\nNenhuma linha correspondente ao padrão da tabela foi encontrada no texto.")
    print("Verifique o texto bruto extraído abaixo para ajustar o padrão se necessário:")
    print("-" * 50)
    print(texto_bruto)
    print("-" * 50)