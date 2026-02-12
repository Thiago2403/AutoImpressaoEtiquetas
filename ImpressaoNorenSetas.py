import xlwings as wx
import time 
from win32com.client import constants

intervalo = 10

arquivo = "" #Caminho do arquivo

app = wx.App(visible = True) # Deixa o app aberto para consultas 
app.display_alerts = False # 

wb = app.books.open(arquivo) # Abre o arquivo
sheet = wb.sheets.active # pega a planilha ativa no momento e guarda na variavel sheet
rng = sheet.range("J7") # Pega o conteudo da célula J7

def definir_direcao(rua,acao):
    rua = rua.upper()
    acao = acao.upper()

    if rua in ["A","C","E"]:
        return "RIGHT" if acao == "COLETAR" else "LEFT"

    if rua in ["B","D","F"]:
        return "LEFT" if acao == "COLETAR" else "RIGHT"

    return None

def ajustar_seta(shape,direcao,texto):
    if direcao == "RIGHT":
        shape.api.AutoShapeType = constants.msoShapeRightArrow
    else:
        shape.api.AutoShapeType = constants.msoShapeLeftArrow

    shape.api.TextFrame2.TextRange.Text = texto

def atualizar_etiquetas(sheet, prefixo , quantidade, direcao):
    for i in range(1, quantidade + 1):
        shape = sheet.shapes(f"{prefixo}_Seta{i}")
        ajustar_seta(shape,direcao,str(i))
    
    shape_fim = sheet.shapes(f"{prefixo}_SetaFim")
    ajustar_seta(shape_fim, direcao, "FIM")

print("Impressao Iniciada! Aperte CTRL + C para parar")

try:
    while True : 
        rua = sheet.range("E16").value
        acao = sheet.range("B18").value
        direcao = definir_direcao(rua, acao)

        maguchi1 = int(sheet.range("F3").value)  # NS1
        contador1 = int(rng.value) 

        maguchi2 = int(sheet.range("F8").value)  # NS2
        contador2 = int(rng.value) 

        terminou_ns1 = False
        terminou_ns2 = False


        # -------- NS1 --------
        if contador1 > maguchi1:
            terminou_ns1 = True
        else:
            if contador1 == maguchi1:
                texto_ns1 = f"{contador1} FIM"
            else:
                texto_ns1 = str(contador1)

            shape_ns1 = sheet.shapes("NS1_Seta")
            ajustar_seta(shape_ns1, direcao, texto_ns1)
            sheet.range("J7").value += 1

        # -------- NS2 --------
        if contador2 > maguchi2:
            terminou_ns2 = True
        else:
            if contador2 == maguchi2:
                texto_ns2 = f"{contador2} FIM"
            else:
                texto_ns2 = str(contador2)

            shape_ns2 = sheet.shapes("NS2_Seta")
            ajustar_seta(shape_ns2, direcao, texto_ns2)
            sheet.range("J8").value += 1

        # -------- FINAL --------
        if terminou_ns1 and terminou_ns2:
            print("Todas as etiquetas impressas")
            break

        print("Impressão das Etiquetas Concluída!")
        # sheet.api.PrintOut()

        time.sleep(intervalo)

except KeyboardInterrupt:
    print("Programa Parado pelo Usuario")       #popop
