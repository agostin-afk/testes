'   Esse codigo serve para enviar dados de uma serie de colunas para uma planilha 'dados' e com isso organizar para posteriormente plotar graficos com os valores, um de cada celula

Sub Enviar_dados()
'
' Enviar_dados Macro
'
'
    Dim preencher_lista As Integer
    Dim x As Integer
    Dim ws As Worksheets
    Dim lista(1 To 9) As Variant
    For preencher_lista = 1 To 9
        If Sheets("exercícios").Cells(40, preencher_lista) > 0 Then
            lista(preencher_lista) = Sheets("exercícios").Cells(40, preencher_lista).Value
        Else
            lista(preencher_lista) = Null
        End If
    Next preencher_lista
    For x = 1 To 9
        If lista(x) > 0 Then
        Sheets("dados").Cells(Sheets("dados").Cells(1, 26 + x).Value, x).Value = lista(x)
        Sheets("dados").Cells(1, 26 + x).Value = Sheets("dados").Cells(1, 26 + x).Value + 1
        End If
    Next x
    Sheets("exercícios").Range("B2:B15").Value = Null
    Sheets("exercícios").Range("F2:F15").Value = Null
    Sheets("exercícios").Range("J2:J15").Value = Null
    Sheets("exercícios").Range("N2:N15").Value = Null
    Sheets("exercícios").Range("R2:R15").Value = Null
    Sheets("exercícios").Range("D21:D34").Value = Null
    Sheets("exercícios").Range("H21:H34").Value = Null
    Sheets("exercícios").Range("L21:L34").Value = Null
    Sheets("exercícios").Range("P21:P34").Value = Null   
End Sub