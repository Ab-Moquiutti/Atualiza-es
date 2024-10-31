import pandas
import win32com.client
import schedule
import time

def executar_macro():
    arquivo_excel = r'\\magneto\data_file_system\MIS\_CLIENTES\PICPAY\PicPay - Analitico Retencao Churn.xlsb'
    macro_nome = 'RunProcess'  

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True

    workbook = excel.Workbooks.Open(arquivo_excel)
    excel.Application.Run(macro_nome)

    workbook.Save()
    workbook.Close(SaveChanges=True)

    excel.Application.Quit()

    print("Macro executada e planilha atualizada com sucesso! Vá e valide os resultados!")

# Agenda a tarefa para ser executada todos os dias às 13:00
schedule.every().day.at("13:00").do(executar_macro)

# Loop para manter o script em execução
while True:
    schedule.run_pending()
    time.sleep(1)
