Option Explicit

Sub verifyInfo()

    'A função verifca se há informação inserida no local certo.
    Dim sheet As String
    Dim cell As String

    'Adicionar nome da "Aba"
    'Adicionar a "Celula" desejada.
    
    '--Inputs
    sheet = "Planilha1"
    cell = "A1"

        If Worksheets(sheet).range(cell).Value = "" Then
            MsgBox "Por favor, insira a informação necessária!"
            End
        End If

End Sub