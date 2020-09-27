Option Explicit

Sub outlookConnection()

'Conecta Excel ao Outlook
Dim outlookMail As String, folder As String
Dim outApp As Outlook.Application 'Conexão com o Outlook
Dim outMapi As Outlook.MAPIFolder 'Conexão com as pastas desejadas, acesso ao e-mail
Dim outHTML As MSHTML.HTMLDocument 'Variavel das informações dentro do e-mail
Dim outMail As Outlook.MailItem 'Variável do objeto e-mail

' Ativar referências bibliotecas:
' VBA aberto vá em "Ferramentas" -> "Preferências" -> 
' "Microsoft Outlook 16.0 Object library" e "Microsoft HTML Object Library"

'--Input
outlookMail = "example@example.com"
folder = "Caixa de Entrada"

On Error Resume Next
    Set outApp = GetObject(, "OUTLOOK.APPLICATION") 'Tenta configurar a aplicação do outlook
        If (outApp Is Nothing) Then 'Se outlook não estiver aberto
            Set outApp = CreateObject("OUTLOOK.APPLICATION") 'Cria, iniciando e configurando a aplicação Outlook
        End If
On Error GoTo 0

Set outMapi = outApp.GetNamespace("MAPI").Folders(outlookMail).Folders(folder)
Set outHTML = New MSHTML.HTMLDocument 'Configura a variável HTML document para ler o corpo do e-mail

End Sub