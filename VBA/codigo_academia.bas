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
        End If
    Next preencher_lista
    For x = 1 To 9
        If lista(x) <> 0 Then
        Sheets("dados").Cells(Cells(1, 26 + x), x).Value = lista(x)
        Sheets("dados").Cells(1, 26 + x).Value = Sheets("dados").Cells(1, 26 + x).Value + 1
        End If
    Next x
End Sub

