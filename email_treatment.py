from pathlib import Path # modulo importante do python
import win32com.client # pip install pywin32
import os
import zipfile

output_dir = Path.cwd() / "Output" # criando um diretorio no mesmo caminho do script com o nome de Output
output_dir.mkdir(parents=True, exist_ok=True) # parametros parents= true para caso as pastas pais nao existam, entao serao criadas. porem se ja existirem e para nao retornar um erro o existes_ok ira ignorar o erro

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") # basicamente fazendo uma conexao com a API de mensagens da Microsoft Outlook
inbox = outlook.Folders.Item("mary.castro@csu.com.br").Folders.Item("Caixa de Entrada").Folders.Item("CSU_BKO_Pendentes")

# 2
def zip_processing():
    with zipfile.ZipFile(arquivo_zip, 'r') as ref: # utilizando a biblioteca zipfile.Zipfile e 'r' para abrir o arquivo zipado em modo de leitura
        for file in ref.namelist(): # namelist é um metodo proprio da biblioteca zipfile, especial para arquivos zipados
            ref.extract(file, output_dir) # extraindo os arquivos do zip para a pasta especificada:: extrair(o arquivo, para a pasta especificada)
# 1
# itera sobre todos os emails da pasta escolhida
for item in inbox.Items:
    if item.Attachments.Count > 0:
        received_time = item.ReceivedTime.strftime('%Y-%m-%d %H:%M:%S') # armazenando o horario dos emails que contém anexos
        for attachments in item.Attachments:
            if attachments.FileName.endswith(".zip"):
                arquivo_zip = os.path.join(output_dir, attachments.FileName) # acessando a pasta que será salva temporariamente os arquivos
                attachments.SaveAsFile(arquivo_zip)
