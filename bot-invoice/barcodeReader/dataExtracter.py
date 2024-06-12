import PyPDF2
import re

path_pdf = r'C:\Users\leoro\Desktop\Projetos\DMStores\Automatização de Devoluções\barcodeReader\boleto.pdf'


def data_extracter_venc(path_pdf):
    with open(path_pdf, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)

        texto = pdf_reader.pages[0].extract_text()
        padrao_data = r'\d{2}/\d{2}/\d{4}'
        datas_vencimento = re.findall(padrao_data, texto)
        if datas_vencimento:
            return datas_vencimento[1]
        else:
            return None


def data_extracter_valor(path_pdf):
    with open(path_pdf, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)

        texto = pdf_reader.pages[0].extract_text()
        padrao_valor = r'Valor\s+do\s+Documento\s*R\$\s*([\d,.]+)'
        valores = re.findall(padrao_valor, texto)

        if valores:
            return valores[0]
        else:
            return None