import xlwings as wx
import time 
from win32com.client import constants

intervalo = 10

arquivo = "" #Caminho do arquivo.

app = wx.App(visible = True) # Deixa o app aberto para consultas .
app.display_alerts = False # 

wb = app.books.open(arquivo) # Abre o arquivo.
sheet = wb.sheets.active # pega a planilha ativa no momento e guarda na variavel sheet
rng = sheet.range("J7") # Pega o conteudo da célula J7.

def definir_direcao(rua,acao):.
    rua = rua.upper() # Deixar maiúsculo
    acao = acao.upper() # Deixar maiúsculo.

    if rua in ["A","C","E"]:
        return "RIGHT" if acao == "COLETAR" else "LEFT" # Retorna Right se a célula nas Ruas [A,C,E] forem Coletar, se for abastecer retorna LEFT..

    if rua in ["B","D","F"]:
        return "LEFT" if acao == "COLETAR" else "RIGHT" # Retorna LEFT se a célula nas Ruas [B,D,F] forem Coletar, se for Abastecer retorna RIGHT.

    return None # Função sem retorno 

def ajustar_seta(shape,direcao,texto):
    if direcao == "RIGHT":
        shape.api.AutoShapeType = constants.msoShapeRightArrow # Ajusta a seta para direita se a direção = RIGHT.
    else:
        shape.api.AutoShapeType = constants.msoShapeLeftArrow # Ajusta a seta para a esquerda se a direção = LEFT .

    shape.api.TextFrame2.TextRange.Text = texto # Atualiza o texto pelo conteúdo da variável texto.

def atualizar_etiquetas(sheet, prefixo , quantidade, direcao):
    for i in range(1, quantidade + 1):
        shape = sheet.shapes(f"{prefixo}_Seta{i}")# Busca a seta na planilha -----------------------------     Arrumar
        ajustar_seta(shape,direcao,str(i)) # Ajusta a seta
    

print("Impressao Iniciada! Aperte CTRL + C para parar")

try:
    while True : 
        rua = sheet.range("E16").value # Célula do endereço
        rua2 = sheet.range("E42").value # Célula de endereço do segundo numero de série

        acao1_1 = sheet.range("B18").value # Célula de Coleta ou Abastecer da 1° etiqueta
        acao1_2 = sheet.range("B31").value # Célula de Coleta ou Abastecer da 2° etiqueta
        acao2_1 = sheet.range("B44").value # Célula de Coleta ou Abastecer da 3° etiqueta
        acao2_2 = sheet.range("B57").value # Célula de Coleta ou Abastecer da 4° etiqueta

        direcao1_1 = definir_direcao(rua, acao1_1)# Define a direção da seta 1 
        direcao1_2 = definir_direcao(rua, acao1_2)# Define a direção da seta 2
        direcao2_1 = definir_direcao(rua2, acao2_1)# Define a direçao da seta 3
        direcao2_2 = definir_direcao(rua2, acao2_2)# Define a direçãp da seta 4 

        maguchi1 = int(sheet.range("F3").value)  # NS1
        contador1 = int(rng.value) 

        maguchi2 = int(sheet.range("F8").value)  # NS2
        contador2 = int(rng.value) 

        terminou_ns1 = False
        terminou_ns2 = False


        # -------- NS1 --------
        if contador1 > maguchi1:
            terminou_ns1 = True # Contador para parar os maguchis
        else:
            if contador1 == maguchi1:
                texto_ns1 = f"{contador1} FIM" # Adiciona a palavra 'FIM' na ultima seta 
            else:
                texto_ns1 = str(contador1) # Numero do maguchi

            shape1_1 = sheet.shapes("Seta1_1") # Definindo a seta 1
            shape1_2 = sheet.shapes("Seta2_2") # Definindo a seta 2
            ajustar_seta(shape1_1, direcao1_1, texto_ns1) # Ajustando a direçao da seta 1
            ajustar_seta(shape1_2, direcao1_2, texto_ns1) # Ajustando a direção da seta 2
            sheet.range("N7").value += 1 # Somando o contador de maguchis

        # -------- NS2 --------
        if contador2 > maguchi2:
            terminou_ns2 = True # Contador para parar os maguchis
        else:
            if contador2 == maguchi2:
                texto_ns2 = f"{contador2} FIM" # Adiciona a palavra 'FIM' na ultima seta 
            else:
                texto_ns2 = str(contador2) # Numero do maguchi

            shape2_1 = sheet.shapes("Seta2_1") # Ajustando a direçao da seta 3
            shape2_2 = sheet.shapes("Seta2_2") # Ajustando a direção da seta 4
            ajustar_seta(shape2_1, direcao2_1, texto_ns2) # Ajustando a direçao da seta 3
            ajustar_seta(shape2_2, direcao2_2, texto_ns2) # Ajustando a direçao da seta 4
            sheet.range("Q7").value += 1

        print("Impressão das Etiquetas Concluída!")
        # sheet.api.PrintOut()

        # -------- FINAL --------
        if terminou_ns1 and terminou_ns2:
            print("Maguchis finalizados - avançando o numero de serie")
            
            sheet.range("J7").value += 1


        time.sleep(intervalo)

except KeyboardInterrupt:
    print("Programa Parado pelo Usuario")       #popop
