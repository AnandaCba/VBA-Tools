Option Explicit

Sub verifyRows()
    
    'Verifica a quantidade de linhas no inicio do VBA e no final.
    'É necessário chamar a função no inicio do vba e no final.
    Dim sheet As String
    Dim cRows As Long
    Dim column As String
    Dim first As Long
    Dim last As Long
    

    'insira o nome da "aba" da planilha que deseja fazer a verificação.
    'Insira a coluna que sará a verificação
    
    '--Inputs
    sheet = "Planilha1"
    column = "A"

        If first = 0 Then
            cRows = 1
                While Worksheets(sheet).Cells(cRows, column).Value <> ""
                    cRows = cRows + 1
                Wend
            cRows = cRows - 1
            first = cRows
        Else
            cRows = 1
                On Error Resume Next
                    While Worksheets(sheet).Cells(cRows, column).Value <> ""
                        cRows = cRows + 1
                    Wend
                On Error GoTo 0
            cRows = cRows - 1
            last = cRows
            MsgBox ("Inicio com: " & first & " Linhas." & Chr(13) & "Final com: " & last & " Linhas.")
        End If

End Sub