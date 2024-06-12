from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from datetime import datetime
import os


def generate_report(lista_pedidos, lista_valores):

    if len(lista_pedidos) != len(lista_valores):
        print("Não consegui gerar o relatório, as listas tem tamanhos diferentes.")
        return

    total_valor = sum(float(valor.replace(',', '.')) for valor in lista_valores)

    total_pedidos = len(lista_pedidos)

    print()
    print(f"Paguei {total_pedidos} boletos, com o valor total de R$ {total_valor}")

    titulo = "Relatório de Taxações - "
    data_atual = datetime.now().strftime("%d/%m/%Y")
    nome_arquivo = f"Relatorio_Taxacoes_{total_valor}.pdf"
    diretorio = r"C:\Users\leoro\Desktop\Projetos\DMStores\Automatização de Devoluções\barcodeReader\relatorios"
    caminho_arquivo = os.path.join(diretorio, nome_arquivo)

    texto = titulo + data_atual

    c = canvas.Canvas(caminho_arquivo, pagesize=letter)

    c.setFont("Helvetica-Bold", 20)
    c.drawString(100, 750, texto)

    c.setFont("Helvetica-Bold", 16)
    y = 700

    c.drawString(100, y, f"Total de Boletos Pagos: {total_pedidos}")
    y -= 30
    c.drawString(100, y, f"Total do Valor Pago: R$ {total_valor}")
    y -= 50

    c.setFont("Helvetica", 12)

    for pedido, valor in zip(lista_pedidos, lista_valores):
        text = f"Pedido: {pedido}           -           Valor: R$ {valor}"
        c.drawString(100, y, text)
        y -= 20

    c.save()
    print()
    print("Relatório gerado, conferir pasta!")


