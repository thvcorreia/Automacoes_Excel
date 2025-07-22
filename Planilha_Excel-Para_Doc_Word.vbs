Sub PreencherWordComDados()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim caminhoModelo As String
    Dim nomeArquivo As String
    Dim linha As Long
    Dim Loja As String, Nome As String, Email As String, CPF As String, NS As String, Patrimonio As String
    
    On Error GoTo Erro ' Tratamento de erro

    ' Caminho do seu modelo Word
    caminhoModelo = "C:\Users\thiago.correia\Downloads\Teste_VBA\0.Termo_padrao.docx" ' <-- ajuste esse caminho

    ' Cria instância do Word
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True
    wdApp.DisplayAlerts = False

    linha = 2 ' Começa da segunda linha

    Do While Not IsEmpty(Sheets("Planilha1").Cells(linha, 1).Value)

        Loja = CStr(Sheets("Planilha1").Cells(linha, 1).Value)
        Nome = CStr(Sheets("Planilha1").Cells(linha, 2).Value)
        Email = CStr(Sheets("Planilha1").Cells(linha, 3).Value)
        CPF = CStr(Sheets("Planilha1").Cells(linha, 4).Value)
        NS = CStr(Sheets("Planilha1").Cells(linha, 5).Value)
        Patrimonio = CStr(Sheets("Planilha1").Cells(linha, 6).Value)
        

        Set wdDoc = wdApp.Documents.Open(caminhoModelo)
        Application.Wait Now + TimeValue("0:00:01")

        With wdDoc.Content.Find
            .Text = "{{Loja}}"
            .Replacement.Text = Loja
            .Forward = True
            .Wrap = 1
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .Execute Replace:=2
        End With

        With wdDoc.Content.Find
            .Text = "{{Nome}}"
            .Replacement.Text = Nome
            .Execute Replace:=2
        End With

        With wdDoc.Content.Find
            .Text = "{{Email}}"
            .Replacement.Text = Email
            .Execute Replace:=2
        End With
        
        With wdDoc.Content.Find
            .Text = "{{CPF}}"
            .Replacement.Text = CPF
            .Execute Replace:=2
        End With
        
        With wdDoc.Content.Find
            .Text = "{{NS}}"
            .Replacement.Text = NS
            .Execute Replace:=2
        End With
        
        With wdDoc.Content.Find
            .Text = "{{Patrimonio}}"
            .Replacement.Text = Patrimonio
            .Execute Replace:=2
        End With

        nomeArquivo = "C:\Users\thiago.correia\Downloads\Teste_VBA\" & "Termo-Padrao_" & Replace(Nome, " ", "_") & ".docx"
        wdDoc.SaveAs2 nomeArquivo
        wdDoc.Close False

        linha = linha + 1
    Loop

Finalizar:
    wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing
    MsgBox "Documentos gerados com sucesso!"
    Exit Sub

Erro:
    MsgBox "Erro na linha " & linha & ": " & Err.Description
    Resume Finalizar
End Sub


