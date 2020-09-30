Option Explicit

Sub cleanColumns()

    'Remove as colunas indesejadas com base no cabeçalho
    Dim cColumns As Long
    Dim checkArr As Long
    Dim qtdCheckArr As Long
    Dim qtdColumns As Long
    Dim sheet As String
    Dim saveColumn As String
    Dim nameColumn(1 To 5) As String
    
    'Insira a "aba" desejada.
    'Insira a quantidade de colunas a ser analisada
    'Insira a quantidade de colunas que VÃO FICAR dentro do array.
    'Insira os nomes das colunas que VÃO ficar.
    
    '--Inputs
    sheet = "Planilha1"
    qtdColumns = 10
    qtdCheckArr = 5
    
    '--Inputs nome das colunas.
    nameColumn(1) = "INDENIZ"
    nameColumn(2) = "NF"
    nameColumn(3) = "VAL_NF"
    nameColumn(4) = "DESCR_EMPRESA"
    nameColumn(5) = "MODAL"
    
        For cColumns = 1 To qtdColumns
            For checkArr = 1 To qtdCheckArr
                If Worksheets(sheet).Cells(1, cColumns).Value = nameColumn(checkArr) Then
                    saveColumn = nameColumn(checkArr)
                End If
            Next checkArr

                If Worksheets(sheet).Cells(1, cColumns).Value <> saveColumn Then
                    Worksheets(sheet).Columns(cColumns).ClearContents
                End If
        Next cColumns

        For cColumns = qtdColumns To 1 Step -1
            If Worksheets(sheet).Cells(1, cColumns).Value = "" Then
                Columns(cColumns).Delete
            End If
        Next cColumns

End Sub