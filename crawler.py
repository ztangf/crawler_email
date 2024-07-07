import win32com.client as win32
from pathlib import Path

# Criando uma pasta para armazenar os arquivos de anexos
destino = Path.cwd()/"output"
destino.mkdir(parents=True,exist_ok=True)

#Instanciando Outlook
outlook = win32.Dispatch("Outlook.Application").getNamespace("MAPI")

#acessando pasta especifica
root = outlook.Folders.Item(1)

#mostrando pastas 
for folder in root.Folders:
    print(folder.Name)

inbox = outlook.GetDefaultFolder(6)

#pasta específica 
inbox = root.Folders["Caixa de Entrada"].Folders["teste"]

messages = inbox.Items

#Option1 
for m in messages:
    subject = m.Subject
    body = m.Body
    attachment = m.Attachments

    #criando pasta com assunto do email
    pasta_destino = destino / str(subject).replace(':','').replace('/','')
    pasta_destino.mkdir(parents=True, exist_ok=True)

    #criando arquivo com corpo do email
    Path(pasta_destino / 'Corpo_Email.txt').write_text(body)

#Option2 
for message in messages:
    if message.Attachments.Count > 0:
        for attachment in message.Attachments:
            attachment.SaveAsFile(Path.cwd() / "output" / attachment.FileName)