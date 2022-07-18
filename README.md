# enviodeemailpy
codigo para envio de email
#enviar email atraves do outlook instalado no pc

#instalar py win32
import win32com.client as win32

#integrar py ao outlook
outlook = win32.Dispatch ('outlook.application')

#criar email
email = outlook.CreateItem(0)

#configuração do email

email.To = 'destinatario'
email.Subject = 'assunto'
email.HTMLBody = '''corpo do email''' #paragrafo em html <p> </p>


#anexar arquivos no email
anexo = 'local do arquivo'
email.Attachments.Add(anexo) #repetir variavel para mais de um arquivo


email.Send
print('E-mail Enviado')
