import win32com.client as win32
from pathlib import Path
import re

outlook = win32.Dispatch('Outlook.Application').GetNamespace('MAPI')    #Abre o outlook
root_folders = outlook.Folders.item(3)                                  #Seleciona o e-mail que irei utilizar, como possuo 3 e-mail dentro do outlook, peguei o ultimo de cima para baixo.
inbox = root_folders.Folders['Errado']                                  #Seleciona a pasta dentro do e-mail.
messages = inbox.items                                                  #Seleciona as mensagens dentro da pasta.

f = open('texto.txt', 'r+')     #Abre o arquivo de texto previamente criado.
f.truncate(10)                  #Define um tamanho para o arquivo de texto.

rest = [] #Cria uma lista vazia para ser usada na comparação de strings.

for m in messages:
    subject = m.Subject             #Assunto do e-mail
    body = m.body                   #Corpo do e-mail
    attachments = m.Attachments     #Anexos do e-mail
    leitura = m.Unread              #E-mail não lido

    if leitura:
        lista = re.findall(r'[a-zA-z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+', body) #Busca um padrão dentro do corpo do e-mail.
        printer = ', '.join(lista)                                                  #Converte a lista em string. Já que o re.findall retorna uma lista.
        f.writelines(printer)                                                       #Escreve a string no arquivo de texto.
        f.write('\n')                                                               #Pula uma linha no arquivo de texto.
        rest.append(printer)                                                        #Adiciona a string na lista REST, para comparar fora do FOR.
        
diferentes = set(rest)  #Exclui as strings repetidas.

f.truncate(0)   #Apagada o arquivo de texto com os e-mails repetidos.
f.seek(0)       #Volta ao inicio do arquivo de texto.

for lines in diferentes:
    f.writelines(lines) #Cria o arquivo de texto baseado na variavel diferentes. Assim não ficam e-mails repetidos dentro do arquivo.
    f.write('\n')       #Pula uma linha no arquivo de texto.

f.close()   #Encerra o arquivo