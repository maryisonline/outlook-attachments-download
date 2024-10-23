from pathlib import Path # modulo importante do python
import win32com.client # pip install pywin32
import os
import zipfile
import pandas as pd

output_dir = Path.cwd() / "Output" # criando um diretorio no mesmo caminho do script com o nome de Output
output_dir.mkdir(parents=True, exist_ok=True) # parametros parents= true para caso as pastas pais nao existam, entao serao criadas. porem se ja existirem e para nao retornar um erro o existes_ok ira ignorar o erro

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") # basicamente fazendo uma conexao com a API de mensagens da Microsoft Outlook
inbox = outlook.Folders.Item("your@email.com.br").Folders.Item("Caixa de Entrada").Folders.Item("Pasta_Desejada")

# 2
def zip_processing(arquivo_zip):
    with zipfile.ZipFile(arquivo_zip, 'r') as ref: # utilizando a biblioteca zipfile.Zipfile e 'r' para abrir o arquivo zipado em modo de leitura
        for file in ref.namelist(): # namelist é um metodo proprio da biblioteca zipfile, especial para arquivos zipados
            ref.extract(file, output_dir) # extraindo os arquivos .csv do zip para a pasta especificada:: extrair(o arquivo, para a pasta especificada)

# 1
# itera sobre todos os emails da pasta escolhida
for item in inbox.Items:
    if item.Attachments.Count > 0:
        received_time = item.ReceivedTime.strftime('%H:%M:%S') # armazenando o horario dos emails que contém anexos
        received_date = item.ReceivedTime.strftime('%Y-%m-%d') # armazenando a data dos emails que contém anexos
        for attachments in item.Attachments:
            if attachments.FileName.endswith(".zip"):
                arquivo_zip = os.path.join(output_dir, attachments.FileName) # acessando a pasta que será salva temporariamente os arquivos
                attachments.SaveAsFile(arquivo_zip)
                zip_processing(arquivo_zip) # chama a funcao com o parametro
                # os.remove(arquivo_zip) # apaga o arquivo zipado
                # print(f'baixo o {arquivo_zip} eeee')

def concatena():
    for file in os.listdir(output_dir):
        file_to_read = os.path.join(output_dir, file)
        df = pd.read_csv(file_to_read, encoding='ISO-8859-1') # lendo os arquivos que sobraram (csv)
        df.insert(0, 'Hora Recebida', received_time)
        df.insert(1, 'Data Recebida', received_date)                    
        if os.path.exists(arq_csv):
            df_consolidado = pd.read_csv(arq_csv, sep=';', encoding='ISO-8859-1') # "le" o arquivo
            df_consolidado = pd.concat([df_consolidado, df], ignore_index=True, axis=0) # concatena no arquivo consolidado
        else:
            df_consolidado = df # se o arquivo consolidado nao existir, cria um a partir do
        df_consolidado.to_csv(arq_csv, index=False, sep=';', encoding='ISO-8859-1') # salva o arquivo
