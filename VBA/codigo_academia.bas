Function exercicios(nome As String) As Collection
    Dim grupo_cels As Range
    Dim tonalidade_cels As Range
    Dim resultado As New Collection
    Select Case nome
        Case "peito"
            Set grupo_cels = Sheets("exercícios").Range(Sheets("exercícios").Cells(2, 2), Sheets("exercícios").Cells(15, 2))
            Set tonalidade_cels = Sheets("dados").Range(Sheets("dados").Cells(1, 53), Sheets("dados").Cells(14, 53))
        Case "biceps"
            Set grupo_cels = Sheets("exercícios").Range(Sheets("exercícios").Cells(2, 6), Sheets("exercícios").Cells(15, 6))
            Set tonalidade_cels = Sheets("dados").Range(Sheets("dados").Cells(1, 54), Sheets("dados").Cells(14, 54))
        Case "posterior_de_coxa"
            Set grupo_cels = Sheets("exercícios").Range(Sheets("exercícios").Cells(2, 10), Sheets("exercícios").Cells(15, 10))
            Set tonalidade_cels = Sheets("dados").Range(Sheets("dados").Cells(1, 55), Sheets("dados").Cells(14, 55))
        Case "ombro"
            Set grupo_cels = Sheets("exercícios").Range(Sheets("exercícios").Cells(2, 14), Sheets("exercícios").Cells(15, 14))
            Set tonalidade_cels = Sheets("dados").Range(Sheets("dados").Cells(1, 56), Sheets("dados").Cells(14, 56))
        Case "triceps"
            Set grupo_cels = Sheets("exercícios").Range(Sheets("exercícios").Cells(2, 18), Sheets("exercícios").Cells(15, 18))
            Set tonalidade_cels = Sheets("dados").Range(Sheets("dados").Cells(1, 57), Sheets("dados").Cells(14, 57))
        Case "costas"
            Set grupo_cels = Sheets("exercícios").Range(Sheets("exercícios").Cells(21, 4), Sheets("exercícios").Cells(34, 4))
            Set tonalidade_cels = Sheets("dados").Range(Sheets("dados").Cells(1, 58), Sheets("dados").Cells(14, 58))
        Case "antebraco"
            Set grupo_cels = Sheets("exercícios").Range(Sheets("exercícios").Cells(21, 8), Sheets("exercícios").Cells(34, 8))
            Set tonalidade_cels = Sheets("dados").Range(Sheets("dados").Cells(1, 59), Sheets("dados").Cells(14, 59))
        Case "quadriceps"
            Set grupo_cels = Sheets("exercícios").Range(Sheets("exercícios").Cells(21, 12), Sheets("exercícios").Cells(34, 12))
            Set tonalidade_cels = Sheets("dados").Range(Sheets("dados").Cells(1, 60), Sheets("dados").Cells(14, 60))
        Case "gluteo"
            Set grupo_cels = Sheets("exercícios").Range(Sheets("exercícios").Cells(21, 16), Sheets("exercícios").Cells(34, 16))
            Set tonalidade_cels = Sheets("dados").Range(Sheets("dados").Cells(1, 61), Sheets("dados").Cells(14, 61))
        Case Else
            Set grupo_cels = Nothing
            Set tonalidade_cels = Nothing
    End Select
    If Not grupo_cels Is Nothing Then
        resultado.Add grupo_cels, "Grupo"
    End If
    If Not tonalidade_cels Is Nothing Then
        resultado.Add tonalidade_cels, "Tonalidade"
    End If
    
    Set exercicios = resultado
End Function

Sub AplicarFormatacao(nomeGrupo As String)
    On Error GoTo ErrorHandler
    Dim cel_ao_lado As Range
    Dim tintAndShadeValue As Double
    Dim resultado As Collection
    Dim tonalidade As Range
    Dim passo As Integer
    Dim cel As Range
    passo = 1
    Set resultado = exercicios(nomeGrupo)
    Set tonalidade = resultado("Tonalidade")
    For Each cel In resultado("Grupo")
        If Not IsEmpty(cel.Value) Then
            Set cel_ao_lado = cel.Offset(0, -1)
            With cel_ao_lado.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent6
                tintAndShadeValue = tonalidade.Cells(passo, 1).Value
                tonalidade.Cells(passo, 1).Value = tonalidade.Cells(passo, 1).Value - 0.043333
                If tintAndShadeValue >= -0.25 And tintAndShadeValue <= 1 Then
                    .TintAndShade = tintAndShadeValue
                Else
                    .Color = RGB(255, 0, 0) ' Valor padrão
                End If
                .PatternTintAndShade = 0
            End With
        End If
        passo = passo + 1
    Next cel
    Exit Sub
ErrorHandler:
    MsgBox "Ocorreu um erro: " & Err.Description, vbCritical
    Resume Next
End Sub

Sub AplicarFormatacaoEmCelsNaoVazias(nomeGrupo As String)
    Dim cel As Range
    Dim resultado As Collection
    Set resultado = exercicios(nomeGrupo)
    For Each cel In resultado("Grupo")
        If Not IsEmpty(cel.Value) Then
            AplicarFormatacao nomeGrupo
        End If
    Next cel
End Sub
Sub Enviar_dados()
'
' Enviar_dados Macro
'
'
    Dim preencher_lista As Integer
    Dim x As Integer
    Dim ws As Worksheets
    Dim lista(1 To 9) As Variant
    Dim cel As Range
    Dim cel_ao_lado As Range
    Dim tintAndShadeValue As Double
    Application.ScreenUpdating = False
    AplicarFormatacaoEmCelsNaoVazias "peito"
    AplicarFormatacaoEmCelsNaoVazias "biceps"
    AplicarFormatacaoEmCelsNaoVazias "posterior_de_coxa"
    AplicarFormatacaoEmCelsNaoVazias "ombro"
    AplicarFormatacaoEmCelsNaoVazias "triceps"
    AplicarFormatacaoEmCelsNaoVazias "costas"
    AplicarFormatacaoEmCelsNaoVazias "antebraco"
    AplicarFormatacaoEmCelsNaoVazias "quadriceps"
    AplicarFormatacaoEmCelsNaoVazias "gluteo"
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
    Sheets("exercícios").Range("B2:B15").ClearContents
    Sheets("exercícios").Range("F2:F15").ClearContents
    Sheets("exercícios").Range("J2:J15").ClearContents
    Sheets("exercícios").Range("N2:N15").ClearContents
    Sheets("exercícios").Range("R2:R15").ClearContents
    Sheets("exercícios").Range("D21:D34").ClearContents
    Sheets("exercícios").Range("H21:H34").ClearContents
    Sheets("exercícios").Range("L21:L34").ClearContents
    Sheets("exercícios").Range("P21:P34").ClearContents
    Application.ScreenUpdating = True
End Sub

