import openpyxl
from datetime import datetime

# Caminho do arquivo Excel
caminho_arquivo = "/home/lfragoso/projetos/dash-burgetXLogComexXComercial/iTRACKER_novo_01_06_v2.xlsx"

def abrir_e_salvar():
    print("🔄 Abrindo a planilha...")
    wb = openpyxl.load_workbook(caminho_arquivo, data_only=False)
    print("🕒 Atualização realizada em:", datetime.now().strftime('%d/%m/%Y %H:%M:%S'))

    print("💾 Salvando planilha...")
    wb.save(caminho_arquivo)
    print("✅ Planilha salva com sucesso!")

if __name__ == "__main__":
    abrir_e_salvar()
