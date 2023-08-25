from pyzbar.pyzbar import decode
from pdf2image import convert_from_path
import os

def linha_digitavel(linha):
    # https://www.bb.com.br/docs/pub/emp/mpe/dwn/PadraoCodigoBarras.pdf
    def modulo10(num):
        soma = 0
        peso = 2
        for c in reversed(num):
            parcial = int(c) * peso
            if parcial > 9:
                s = str(parcial)
                parcial = int(s[0]) + int(s[1])
            soma += parcial
            if peso == 2:
                peso = 1
            else:
                peso = 2

        resto10 = soma % 10
        if resto10 == 0:
            modulo10 = 0
        else:
            modulo10 = 10 - resto10
        return modulo10

    def monta_campo(campo):
        campo_dv = "%s%s" % (campo, modulo10(campo))
        return "%s.%s" % (campo_dv[0:5], campo_dv[5:])
    
    return ' '.join([monta_campo(linha[0:4] + linha[19:24]),
                        monta_campo(linha[24:34]),
                        monta_campo(linha[34:44]),
                        linha[4],
                        linha[5:19]])


def barcode_reader(pdf_path):
    
    pages = convert_from_path(pdf_path)
    img = pages[0]

    detectedBarcodes = decode(img)

    if not detectedBarcodes:
        return False
    else:
        for barcode in detectedBarcodes:
            if barcode.data != "" and barcode.type == 'I25':
                barcode_ = barcode.data.decode('utf-8')
                return barcode_


if __name__ == '__main__':
    pdfs = [i for i in os.listdir() if ".pdf" in i]
    pdf_path = pdfs[0]
    print(barcode_reader(pdf_path))
    print(linha_digitavel(barcode_reader(pdf_path)))
