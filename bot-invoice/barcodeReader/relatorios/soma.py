import os

diretorio_relatorios = r"C:\Users\leoro\Desktop\Projetos\DMStores\Automatização de Devoluções\barcodeReader\relatorios"
soma_total = 0

for nome_arquivo in os.listdir(diretorio_relatorios):
    if nome_arquivo.startswith("Relatorio_Taxacoes_") and nome_arquivo.endswith(".pdf"):
        # Extrair o valor do nome do arquivo (assumindo que está após o último "_" e antes da extensão ".pdf")
        valor_arquivo = float(nome_arquivo.split("_")[-1][:-4])
        soma_total += valor_arquivo

print(f"Soma total dos Relatórios: {soma_total}")