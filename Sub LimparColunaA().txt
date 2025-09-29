Sub LimparColunaA()
    Dim ws As Worksheet
    Dim cel As Range
    Dim textoOriginal As String

    Set ws = ActiveSheet

    For Each cel In ws.Range("A1:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)
        textoOriginal = NormalizarTexto(cel.Value)
        cel.Value = textoOriginal
    Next cel

    MsgBox "Coluna A limpa com sucesso!", vbInformation
End Sub

Function NormalizarTexto(ByVal texto As String) As String
    Dim i As Integer
    Dim resultado As String
    Dim c As String

    ' Converte para minúsculas (opcional)
    texto = LCase(texto)

    ' Remove acentos
    texto = Replace(texto, "á", "a")
    texto = Replace(texto, "à", "a")
    texto = Replace(texto, "ã", "a")
    texto = Replace(texto, "â", "a")
    texto = Replace(texto, "ä", "a")

    texto = Replace(texto, "é", "e")
    texto = Replace(texto, "è", "e")
    texto = Replace(texto, "ê", "e")
    texto = Replace(texto, "ë", "e")

    texto = Replace(texto, "í", "i")
    texto = Replace(texto, "ì", "i")
    texto = Replace(texto, "î", "i")
    texto = Replace(texto, "ï", "i")

    texto = Replace(texto, "ó", "o")
    texto = Replace(texto, "ò", "o")
    texto = Replace(texto, "õ", "o")
    texto = Replace(texto, "ô", "o")
    texto = Replace(texto, "ö", "o")

    texto = Replace(texto, "ú", "u")
    texto = Replace(texto, "ù", "u")
    texto = Replace(texto, "û", "u")
    texto = Replace(texto, "ü", "u")

    texto = Replace(texto, "ç", "c")

    ' Remove caracteres especiais e pontuações
    For i = 1 To Len(texto)
        c = Mid(texto, i, 1)
        If c Like "[A-Za-z0-9 ]" Then
            resultado = resultado & c
        End If
    Next i

    NormalizarTexto = Trim(resultado)
End Function
