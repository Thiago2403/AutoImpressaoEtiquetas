import xlwings as xw
import time 


intervalo = 10

arquivo = r"C:\Users\X\Documents\MacrosExcel\ImpressaoEtiquetas.xlsx"  # caminho do arquivo

app = xw.App(visible=True) # Deixa o excel visivel
app.display_alerts = False# Nao emite alertas do excel

wb = app.books.open(arquivo)# Abre o arquivo
sheet = wb.sheets.active# Pega a planilha ativa no momento e guarda na variavel sheet
rng = sheet.range("J7")#Acessa a celula da planilha

print("Impressao iniciada. CTRL + C para parar")

try:
    while True:
        print("Impressão realizada")#sheet.api.PrintOut() Imprime
        
        # Soma
        if isinstance(rng.value,(int,float)):
            rng.value += 1 
        else:
            print("J7 Invalido, parando o Programa")
            break
        
        time.sleep(intervalo)
        
except KeyboardInterrupt:
    print("Programa Parado pelo Usuario")
    
   