Sub GerarDocumentoComFormatacaoCorreta()
    Dim wdApp As Object
    Dim wdDocFinal As Object
    Dim wdDocModelo As Object
    Dim caminhoModelo As String
    Dim caminhoFinal As String
    Dim textoModelo As String
    Dim linha As Long
    Dim praca As String, loja As String, nomeCompleto As String
    Dim textoPersonalizado As String
    Dim rngInsert As Object
    Dim inicioTexto As Long, fimTexto As Long
    Dim rngBusca As Object

    On Error GoTo Erro

    ' Caminhos
    caminhoModelo = "C:\Users\thiago.correia\OneDrive - Gentil Negócios\Documentos\ModeloAutomacoes\Praca.docx"        ' <- ajuste aqui
    caminhoFinal = "C:\Users\thiago.correia\OneDrive - Gentil Negócios\Documentos\ModeloAutomacoes\Praca-novo.docx" ' <- ajuste aqui

    ' Iniciar Word
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    wdApp.DisplayAlerts = False

    ' Abrir modelo e capturar conteúdo
    Set wdDocModelo = wdApp.Documents.Open(caminhoModelo)
    textoModelo = wdDocModelo.Content.Text
    wdDocModelo.Close False

    ' Criar novo documento final
    Set wdDocFinal = wdApp.Documents.Add

    linha = 2 ' começa da segunda linha da planilha

    Do While Not IsEmpty(Sheets("Planilha1").Cells(linha, 1).Value)
        ' Lê as variáveis da planilha
        praca = CStr(Sheets("Planilha1").Cells(linha, 2).Value)
        loja = CStr(Sheets("Planilha1").Cells(linha, 5).Value)
        nomeCompleto = CStr(Sheets("Planilha1").Cells(linha, 7).Value)

        ' Substitui no texto do modelo
        textoPersonalizado = Replace(textoModelo, "{{Praca}}", praca)
        textoPersonalizado = Replace(textoPersonalizado, "{{Loja}}", loja)
        textoPersonalizado = Replace(textoPersonalizado, "{{NomeCompleto}}", nomeCompleto)

        ' Inserir no final do documento
        Set rngInsert = wdDocFinal.Range
        rngInsert.Collapse Direction:=0 ' wdCollapseEnd
        inicioTexto = rngInsert.End

        rngInsert.InsertAfter textoPersonalizado & vbCrLf & String(40, "-") & vbCrLf & vbCrLf

        ' Define range do novo trecho inserido
        fimTexto = wdDocFinal.Range.End
        Set rngBusca = wdDocFinal.Range(Start:=inicioTexto, End:=fimTexto)

        ' Formatar todas as ocorrências de Praca
        With rngBusca.Find
            .ClearFormatting
            .Text = praca
            .Forward = True
            .Wrap = 0 ' wdFindStop
            Do While .Execute
                rngBusca.Font.Name = "Arial"
                rngBusca.Font.Size = 36
            Loop
        End With

        ' Formatar todas as ocorrências de Loja
        Set rngBusca = wdDocFinal.Range(Start:=inicioTexto, End:=fimTexto)
        With rngBusca.Find
            .ClearFormatting
            .Text = loja
            .Forward = True
            .Wrap = 0
            Do While .Execute
                rngBusca.Font.Name = "Arial"
                rngBusca.Font.Size = 36
                rngBusca.Font.Bold = True
            Loop
        End With

        ' Formatar todas as ocorrências de NomeCompleto
        Set rngBusca = wdDocFinal.Range(Start:=inicioTexto, End:=fimTexto)
        With rngBusca.Find
            .ClearFormatting
            .Text = nomeCompleto
            .Forward = True
            .Wrap = 0
            Do While .Execute
                rngBusca.Font.Name = "Arial"
                rngBusca.Font.Size = 16
                rngBusca.Font.Italic = True
            Loop
        End With

        linha = linha + 1
    Loop

    ' Salvar documento
    wdDocFinal.SaveAs2 caminhoFinal
    wdDocFinal.Close False

Finalizar:
    wdApp.Quit
    Set wdDocFinal = Nothing
    Set wdDocModelo = Nothing
    Set wdApp = Nothing

    MsgBox "Documento gerado com sucesso!"
    Exit Sub

Erro:
    MsgBox "Erro na linha " & linha & ": " & Err.Description
    Resume Finalizar
End Sub


