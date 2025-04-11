import pandas as pd

# Caminho do arquivo original
arquivo_origem = r"C:\Users\leonardo.fragoso\Desktop\Projetos\dash-burgetXLogComexXComercial\Planilha1_extraida.xlsx"

# Caminho do novo arquivo CSV
arquivo_destino = r"C:\Users\leonardo.fragoso\Desktop\Projetos\dash-burgetXLogComexXComercial\Planilha1_filtrada.csv"

# Colunas desejadas
colunas = ["cliente", "DataEmissao", "Empresa", "TipoAtendimento", "tiponotafiscal", "Status"]

# Leitura e extração das colunas especificadas
df = pd.read_excel(arquivo_origem, usecols=colunas)

# Exportação para CSV
df.to_csv(arquivo_destino, index=False, encoding="utf-8")

print(f"Arquivo CSV criado com sucesso em: {arquivo_destino}")
