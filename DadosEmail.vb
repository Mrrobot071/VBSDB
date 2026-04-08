Option Explicit

'=========================================================
' CONFIGURAÇÕES
'=========================================================
Private Const REQUIRE_ATTACHMENTS As Boolean = True   ' True = só e-mails com anexo
Private Const KEEP_UNREAD_STATUS As Boolean = True    ' True = não mexe em UnRead (recomendado)

'=========================================================
' MACRO PRINCIPAL (SEM COLUNA TIPO)
'=========================================================
Public Sub ExtrairNaoLidos()

    Dim olApp As Object, olNs As Object
    Dim olInbox As Object, olTarget As Object
    Dim itms As Object, it As Object, mail As Object

    Dim ws As Worksheet
    Dim linhaAtual As Long
    Dim nomesIgnorados As Variant

    nomesIgnorados = Array("Lucas", "Marcelo", "Francisco", "Erivaldo")

    On Error GoTo TratarErro

    '--- Outlook
    Set olApp = CreateObject("Outlook.Application")
    Set olNs = olApp.GetNamespace("MAPI")
    Set olInbox = olNs.GetDefaultFolder(6) ' Inbox
    Set olTarget = GetFolder(olInbox, "NFE\01.NOTAS")

    If olTarget Is Nothing Then
        MsgBox "Pasta 'NFE\01.NOTAS' não encontrada. Confirme o caminho no Outlook.", vbCritical
        GoTo Finalizar
    End If

    '--- Planilha
    Set ws = ObterOuCriarPlanilha("DadosEmail")
    ws.Cells.Clear

    ' A:J = 10 colunas (REMOVIDA "TIPO (PED/OC)")
    ws.Range("A1:J1").Value = Array( _
        "DATA RECEBIMENTO", "REMETENTE", "PEDIDO/OC", _
        "Nº NOTA FISCAL", "VALOR (ORIGINAL)", "VENCIMENTO", _
        "FORMA DE PAGAMENTO", "COMPRADOR", "FORNECEDOR", "IARA" _
    )

    With ws.Range("A1:J1")
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    linhaAtual = 2

    '--- SOMENTE NÃO LIDOS
    Set itms = olTarget.items.Restrict("[UnRead] = True")
    itms.Sort "[ReceivedTime]", True

    For Each it In itms

        If TypeName(it) = "MailItem" Then
            Set mail = it

            If mail.UnRead = True Then
                If (Not REQUIRE_ATTACHMENTS) Or (mail.Attachments.Count > 0) Then

                    Dim senderName As String
                    senderName = SafeString(mail.senderName)
                    If IsInArray(senderName, nomesIgnorados) Then GoTo Proximo

                    Dim corpoRaw As String, corpoLinhas As String, corpoCompacto As String
                    Dim assunto As String

                    corpoRaw = SafeString(mail.Body)
                    assunto = SafeString(mail.Subject)

                    If Len(corpoRaw) = 0 And Len(assunto) = 0 Then GoTo Proximo

                    corpoLinhas = NormalizeForLines(corpoRaw)
                    corpoCompacto = RemoveAllWhitespace(corpoRaw)

                    '=============================
                    ' PEDIDO / OC (sem coluna tipo)
                    '=============================
                    Dim pedido As String, oc As String
                    Dim codigoFinal As String

                    pedido = ExtrairPedidoOuOC(corpoLinhas, corpoCompacto, assunto, "PEDIDO")
                    oc = ExtrairPedidoOuOC(corpoLinhas, corpoCompacto, assunto, "OC")

                    If EhPedidoOcValido(pedido) Then
                        codigoFinal = pedido
                    ElseIf EhPedidoOcValido(oc) Then
                        codigoFinal = oc
                    Else
                        GoTo Proximo
                    End If

                    '=============================
                    ' NF
                    '=============================
                    Dim notaFiscal As String
                    notaFiscal = ExtrairNF(corpoLinhas, corpoCompacto, assunto)

                    '=============================
                    ' VALOR (ORIGINAL, texto idêntico)
                    '=============================
                    Dim valorOriginal As String
                    valorOriginal = ExtrairValorOriginal(corpoRaw, corpoLinhas, corpoCompacto)

                    '=============================
                    ' VENCIMENTO
                    '=============================
                    Dim vencTxt As String, vencData As Variant
                    vencTxt = ExtrairVencimento(corpoLinhas, corpoCompacto)
                    vencData = Empty
                    If Len(vencTxt) > 0 Then vencData = ToDateBR(vencTxt)

                    '=============================
                    ' Outros campos
                    '=============================
                    Dim formaPgto As String, comprador As String, fornecedor As String
                    formaPgto = RegexGetLineValue(corpoLinhas, "Forma\s*de\s*Pagamento")
                    comprador = RegexGetLineValue(corpoLinhas, "Comprador")
                    fornecedor = RegexGetLineValue(corpoLinhas, "Fornecedor")

                    '=============================
                    ' GRAVAR NA PLANILHA (A:J)
                    '=============================
                    ws.Cells(linhaAtual, 1).Value = mail.ReceivedTime
                    ws.Cells(linhaAtual, 1).NumberFormat = "dd/mm/yyyy hh:mm"

                    ws.Cells(linhaAtual, 2).Value = senderName
                    ws.Cells(linhaAtual, 3).Value = codigoFinal

                    If Len(notaFiscal) > 0 Then ws.Cells(linhaAtual, 4).Value = notaFiscal

                    With ws.Cells(linhaAtual, 5)
                        .NumberFormat = "@"
                        If Len(valorOriginal) > 0 Then
                            .Value = valorOriginal
                        Else
                            .Value = "NÃO ENCONTRADO"
                        End If
                    End With

                    If IsDate(vencData) Then
                        ws.Cells(linhaAtual, 6).Value = CDate(vencData)
                        ws.Cells(linhaAtual, 6).NumberFormat = "dd/mm/yyyy"
                    ElseIf Len(vencTxt) > 0 Then
                        ws.Cells(linhaAtual, 6).NumberFormat = "@"
                        ws.Cells(linhaAtual, 6).Value = vencTxt
                    End If

                    If Len(formaPgto) > 0 Then ws.Cells(linhaAtual, 7).Value = formaPgto
                    If Len(comprador) > 0 Then ws.Cells(linhaAtual, 8).Value = comprador
                    If Len(fornecedor) > 0 Then ws.Cells(linhaAtual, 9).Value = fornecedor

                    linhaAtual = linhaAtual + 1

                    ' NÃO mexe no status
                    If Not KEEP_UNREAD_STATUS Then
                        On Error Resume Next
                        mail.UnRead = False
                        mail.Save
                        On Error GoTo TratarErro
                    End If

                End If
            End If
        End If

Proximo:
    Next it

    '=============================
    ' PROCX (IARA) - Coluna J
    '=============================
    If linhaAtual > 2 Then
        With ws
            .Cells(1, 10).Value = "IARA"
            .Range("J2").FormulaLocal = _
                "=PROCX(C2;" & _
                "'[Planilha de Chamados IARA+ 2026 v8 - Semestre 1.xlsx]Preechimento'!$O:$O;" & _
                "'[Planilha de Chamados IARA+ 2026 v8 - Semestre 1.xlsx]Preechimento'!$X:$X;" & _
                """AUSENTE""" & _
                ")"
            .Range("J2").AutoFill Destination:=.Range("J2:J" & (linhaAtual - 1))
            .Columns("J").NumberFormat = "@"
        End With
    End If

    '--- Formatação final
    With ws.UsedRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ws.Columns("A:J").AutoFit
    ws.Range("A1:J1").AutoFilter

Finalizar:
    On Error Resume Next
    Set mail = Nothing
    Set it = Nothing
    Set itms = Nothing
    Set olTarget = Nothing
    Set olInbox = Nothing
    Set olNs = Nothing
    Set olApp = Nothing
    On Error GoTo 0
    Exit Sub

TratarErro:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical
    Resume Finalizar

End Sub

'=========================================================
' ======= FUNÇÕES (MANTIDAS IGUAIS AO QUE JÁ TE PASSEI) ====
'=========================================================
Private Function ExtrairPedidoOuOC(ByVal corpoLinhas As String, ByVal corpoCompacto As String, ByVal assunto As String, ByVal label As String) As String
    Dim v As String

    v = OnlyDigits(RegexGetFirstGroup(corpoLinhas, label & "\s*(?:n[º°o]?\s*)?[:\-–—]?\s*([0-9\.\-\/ ]+)"))
    If EhPedidoOcValido(v) Then ExtrairPedidoOuOC = v: Exit Function

    v = OnlyDigits(RegexGetFirstGroup(corpoCompacto, label & "(?:Nº|N|NO)?[:\-–—]*([0-9\.\-\/]+)"))
    If EhPedidoOcValido(v) Then ExtrairPedidoOuOC = v: Exit Function

    v = OnlyDigits(RegexGetFirstGroup(assunto, label & "\s*[:\-–—]?\s*([0-9\.\-\/ ]+)"))
    If EhPedidoOcValido(v) Then ExtrairPedidoOuOC = v: Exit Function

    v = RegexGetFirstGroup(assunto, "(45\d{8})")
    If EhPedidoOcValido(v) Then ExtrairPedidoOuOC = v: Exit Function

    v = RegexGetFirstGroup(corpoCompacto, "(45\d{8})")
    If EhPedidoOcValido(v) Then
        ExtrairPedidoOuOC = v
    Else
        ExtrairPedidoOuOC = ""
    End If
End Function

Private Function ExtrairNF(ByVal corpoLinhas As String, ByVal corpoCompacto As String, ByVal assunto As String) As String
    Dim nf As String

    nf = RegexGetFirstGroup(corpoLinhas, "(?:N[º°o]?\s*(?:da\s*)?Nota\s*Fiscal|Nota\s*Fiscal|N[º°o]?\s*NF(?:-?e)?|NF-?e)\s*[:\-–—]?\s*(\d[\dA-Za-z\.\-\/]*)")
    If Len(nf) = 0 Then
        nf = RegexGetFirstGroup(corpoCompacto, "(?:N[º°o]?NOTAFISCAL|NOTAFISCAL|NF-?E|NF)\s*[:\-–—]*\s*(\d[\dA-Za-z\.\-\/]*)")
    End If
    If Len(nf) = 0 Then
        nf = RegexGetFirstGroup(assunto, "(?:NF|NF-?e)\s*[:\-–—]?\s*(\d[\dA-Za-z\.\-\/]*)")
    End If

    ExtrairNF = Trim$(nf)
End Function

Private Function ExtrairValorOriginal(ByVal corpoRaw As String, ByVal corpoLinhas As String, ByVal corpoCompacto As String) As String
    Dim v As String

    v = RegexGetFirstGroup(corpoLinhas, "(?:Valor|Vlr)\s*[:\-–—]?\s*(R\$\s*\d{1,3}(?:\.\d{3})*,\d{2})")
    If Len(v) > 0 Then ExtrairValorOriginal = Trim$(v): Exit Function

    v = RegexGetFirstGroup(corpoLinhas, "(?:Valor|Vlr)\s*[:\-–—]?\s*(\d{1,3}(?:\.\d{3})*,\d{2})")
    If Len(v) > 0 Then ExtrairValorOriginal = Trim$(v): Exit Function

    v = RegexGetFirstGroup(corpoRaw, "(R\$\s*\d{1,3}(?:\.\d{3})*,\d{2})")
    If Len(v) = 0 Then v = RegexGetFirstGroup(corpoRaw, "(\d{1,3}(?:\.\d{3})*,\d{2})")

    ExtrairValorOriginal = Trim$(v)
End Function

Private Function ExtrairVencimento(ByVal corpoLinhas As String, ByVal corpoCompacto As String) As String
    Dim v As String

    v = RegexGetFirstGroup(corpoLinhas, "Vencimento\s*[:\-–—]?\s*([0-3]?\d[\/\-\.\s][01]?\d[\/\-\.\s]\d{2,4})")
    If Len(v) = 0 Then v = RegexGetFirstGroup(corpoCompacto, "VENCIMENTO[:\-–—]*([0-3]?\d[\/\-\.\s][01]?\d[\/\-\.\s]\d{2,4})")

    ExtrairVencimento = Trim$(v)
End Function

Private Function SafeString(ByVal v As Variant) As String
    On Error Resume Next
    If IsNull(v) Or IsEmpty(v) Then SafeString = "" Else SafeString = CStr(v)
    On Error GoTo 0
End Function

Public Function GetFolder(ByVal rootFolder As Object, ByVal folderPath As String) As Object
    Dim arrFolders() As String, folder As Object, i As Long
    If rootFolder Is Nothing Or Len(folderPath) = 0 Then Exit Function

    arrFolders = Split(folderPath, "\")
    Set folder = rootFolder

    For i = LBound(arrFolders) To UBound(arrFolders)
        On Error Resume Next
        Set folder = folder.Folders(arrFolders(i))
        On Error GoTo 0
        If folder Is Nothing Then Exit For
    Next i

    Set GetFolder = folder
End Function

Public Function ObterOuCriarPlanilha(ByVal nome As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nome)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = nome
    End If
    Set ObterOuCriarPlanilha = ws
End Function

Public Function NormalizeForLines(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)
    t = Replace(t, vbTab, " ")
    t = Replace(t, Chr$(160), " ")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    NormalizeForLines = t
End Function

Public Function RemoveAllWhitespace(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, vbCrLf, "")
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Replace(t, vbTab, "")
    t = Replace(t, Chr$(160), "")
    t = Replace(t, " ", "")
    RemoveAllWhitespace = t
End Function

Public Function RegexGetFirstGroup(ByVal texto As String, ByVal pattern As String) As String
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    With re
        .pattern = pattern
        .IgnoreCase = True
        .Global = False
        .Multiline = True
    End With

    If re.Test(texto) Then
        Set m = re.Execute(texto)(0)
        If m.SubMatches.Count > 0 Then
            RegexGetFirstGroup = Trim$(m.SubMatches(0))
        Else
            RegexGetFirstGroup = Trim$(m.Value)
        End If
    Else
        RegexGetFirstGroup = ""
    End If
End Function

Public Function RegexGetLineValue(ByVal texto As String, ByVal labelPattern As String) As String
    RegexGetLineValue = RegexGetFirstGroup(texto, labelPattern & "\s*[:\-–—]?\s*([^\r\n]+)")
End Function

Public Function OnlyDigits(ByVal s As String) As String
    Dim i As Long, ch As String, r As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then r = r & ch
    Next i
    OnlyDigits = r
End Function

Public Function EhPedidoOcValido(ByVal codigo As String) As Boolean
    EhPedidoOcValido = (Len(codigo) = 10 And Left$(codigo, 2) = "45")
End Function

Public Function ToDateBR(ByVal s As String) As Variant
    Dim t As String, d As Variant
    Dim dia As Integer, mes As Integer, ano As Integer

    t = Trim$(s)
    If t = "" Then ToDateBR = Empty: Exit Function

    t = Replace(t, "-", "/")
    t = Replace(t, ".", "/")
    t = Replace(t, " ", "/")

    Do While InStr(t, "//") > 0
        t = Replace(t, "//", "/")
    Loop

    d = Split(t, "/")
    If UBound(d) <> 2 Then ToDateBR = Empty: Exit Function

    dia = CInt(val(d(0)))
    mes = CInt(val(d(1)))
    ano = CInt(val(d(2)))
    If ano < 100 Then ano = 2000 + ano

    If dia >= 1 And dia <= 31 And mes >= 1 And mes <= 12 And ano >= 1900 And ano <= 2100 Then
        ToDateBR = DateSerial(ano, mes, dia)
    Else
        ToDateBR = Empty
    End If
End Function

Public Function IsInArray(ByVal val As String, ByVal arr As Variant) As Boolean
    Dim element As Variant
    IsInArray = False
    For Each element In arr
        If StrComp(val, element, vbTextCompare) = 0 Then
            IsInArray = True
            Exit Function
        End If
    Next
End Function







