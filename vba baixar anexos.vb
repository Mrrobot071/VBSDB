'vba baixar anexo corrigido extração Nnf2

Option Explicit

' ============================
' Configurações
' ============================
Private Const MARK_AS_READ As Boolean = False   ' True para marcar e-mail como lido após salvar
Private Const FILTRAR_APENAS_PDF_XML As Boolean = True
Private Const LOG_DEBUG As Boolean = True       ' Exibe motivos no Immediate (Ctrl+G)

' Nova opção: pular salvamento se o arquivo (nome final) já existir
Private Const SKIP_IF_EXISTS As Boolean = True

' Ajuste seu caminho padrão
Private Function PASTA_BASE_DESTINO() As String
    PASTA_BASE_DESTINO = Environ$("USERPROFILE") & "\OneDrive - ALLOS\Área de Trabalho\Anexos\"
End Function

' ============================
' Entrada principal
' ============================
Public Sub BaixarAnexosRenomeados()
    On Error GoTo TratamentoErro
    
    Dim olApp As Outlook.Application
    Dim olNamespace As Outlook.NameSpace
    Dim olInbox As Outlook.MAPIFolder
    Dim olSubFolder As Outlook.MAPIFolder
    Dim filteredItems As Outlook.items
    Dim item As Object
    Dim mail As Outlook.mailItem
    Dim att As Outlook.Attachment
    
    Dim pastaDestinoBase As String
    pastaDestinoBase = PASTA_BASE_DESTINO()
    EnsureFolderExistsRecursive pastaDestinoBase
    
    Set olApp = Outlook.Application
    Set olNamespace = olApp.GetNamespace("MAPI")
    Set olInbox = olNamespace.GetDefaultFolder(olFolderInbox)
    
    ' Tenta localizar "NFE > 01.NOTAS" em raiz do mailbox OU dentro da Caixa de Entrada
    Set olSubFolder = TryGetSubFolder(olNamespace, olInbox, "NFE", "01.NOTAS")
    If olSubFolder Is Nothing Then
        MsgBox "Subpasta 'NFE > 01.NOTAS' não encontrada.", vbExclamation, "Aviso"
        GoTo Finalizar
    End If
    
    ' Filtra somente não lidos
    Set filteredItems = olSubFolder.items.Restrict("@SQL=""urn:schemas:httpmail:read"" = 0")
    
    Dim anexosBaixadosNoGeral As Boolean
    anexosBaixadosNoGeral = False
    
    ' Mapa de arquivos salvos nesta execução (evitar duplicidade no mesmo run)
    Dim savedMap As Object
    Set savedMap = CreateObject("Scripting.Dictionary")
    
    Dim corpoRaw As String, corpoLinhas As String, corpoCompacto As String
    Dim assunto As String, senderName As String
    Dim pedido As String, oc As String, pedidoValido As Boolean, ocValido As Boolean
    Dim numeroNF As String, fornecedor As String
    
    For Each item In filteredItems
        If TypeOf item Is Outlook.mailItem Then
            Set mail = item
            If mail.Attachments.Count > 0 Then
                assunto = NzStr(mail.Subject)
                senderName = NzStr(mail.senderName)
                corpoRaw = NzStr(mail.Body)
                If Len(corpoRaw) = 0 Then GoTo ProximoItem
                
                ' Normalizações para extração
                corpoLinhas = NormalizeForLines(corpoRaw)
                corpoCompacto = RemoveAllWhitespace(corpoRaw)
                
                ' Extrai PEDIDO/OC
                pedido = OnlyDigits(RegexGetFirst(corpoLinhas, LabelValuePatternTolerant("PEDIDO")))
                oc = OnlyDigits(RegexGetFirst(corpoLinhas, LabelValuePatternTolerant("OC")))
                If Len(pedido) = 0 Then pedido = OnlyDigits(RegexGetFirst(corpoCompacto, LabelValuePatternAfterCompact("PEDIDO")))
                If Len(oc) = 0 Then oc = OnlyDigits(RegexGetFirst(corpoCompacto, LabelValuePatternAfterCompact("OC")))
                
                pedidoValido = EhPedidoOcValido(pedido)
                ocValido = EhPedidoOcValido(oc)
                
                If Not pedidoValido And Not ocValido Then
                    If LOG_DEBUG Then Debug.Print "Ignorado e-mail sem pedido/OC válido: "; assunto
                    GoTo ProximoItem
                End If
                
                Dim pedidoUsar As String
                pedidoUsar = IIf(pedidoValido, pedido, oc)
                
                ' Extrai NF (se houver; senão tentamos pelo nome do anexo)
                numeroNF = Trim$(RegexGetFirst( _
                    corpoLinhas, _
                    "(?:N[º°o]?\s*(?:da\s*)?Nota\s*Fiscal|Nota\s*Fiscal|N[º°o]?\s*NF(?:-?e)?|NF-?e|NF)\s*[:\-–—]?\s*(\d[\dA-Za-z\.\-\/]*)" _
                ))
                
                ' Extrai fornecedor; fallback = remetente
                fornecedor = RegexGetLineValue(corpoLinhas, "Fornecedor")
                If Len(Trim$(fornecedor)) = 0 Then fornecedor = senderName
                
                ' Prefixo CG x FPP
                Dim prefixo As String
                prefixo = DetectarPrefixo(corpoLinhas, assunto)    ' "FPP" ou "CG"
                
                ' Pasta do pedido (aqui mantém na pasta base; ajuste se quiser por pedido)
                Dim pastaPedido As String
                pastaPedido = pastaDestinoBase
                'pastaPedido = pastaDestinoBase & pedidoUsar & "\"
                EnsureFolderExistsRecursive pastaPedido
                
                Dim i As Long
                Dim salvouNesteEmail As Boolean
                salvouNesteEmail = False
                
                For i = 1 To mail.Attachments.Count
                    Set att = mail.Attachments(i)
                    
                    Dim nomeArq As String, ext As String
                    nomeArq = NzStr(att.fileName)
                    ext = "." & LCase$(ObterExtensao(nomeArq))
                    
                    ' Filtrar extensões
                    If FILTRAR_APENAS_PDF_XML Then
                        If Not (ext = ".xml" Or ext = ".pdf") Then
                            If LOG_DEBUG Then Debug.Print "Ignorado (extensão não permitida): "; nomeArq
                            GoTo ProximoAnexo
                        End If
                    End If
                    
                    ' ---- BLOQUEIO: não baixar se for PO (Purchase Order) pelo NOME ----
                    If IsArquivoPO(LCase$(nomeArq)) Then
                        If LOG_DEBUG Then Debug.Print "Ignorado (PO detectado no nome): "; nomeArq
                        GoTo ProximoAnexo
                    End If
                    
                    ' Detectar tipo (XML / BOLETO / NF / INVALIDO_PO / DESCONHECIDO)
                    Dim tipo As String
                    tipo = DetectarTipoAnexo(nomeArq, ext)
                    
                    If tipo = "INVALIDO_PO" Then
                        If LOG_DEBUG Then Debug.Print "Ignorado (PO detectado pela detecção): "; nomeArq
                        GoTo ProximoAnexo
                    End If
                    If tipo = "DESCONHECIDO" Then
                        If LOG_DEBUG Then Debug.Print "Ignorado (PDF não reconhecido como NF/BOLETO): "; nomeArq
                        GoTo ProximoAnexo
                    End If
                    
                    Dim fornecedorUsar As String
                    fornecedorUsar = SafeFileComponent(fornecedor)
                    
                    ' nfUsar: do corpo ou heurística no arquivo
                    Dim nfUsar As String
                    nfUsar = Trim$(numeroNF)
                    If Len(nfUsar) = 0 Then
                        nfUsar = ExtrairNumeroProvavelDeNF(nomeArq)
                    End If
                    
                    ' ==== NOVO: base de nome única para NF/BOLETO/XML ====
                    Dim baseNome As String
                    If Len(nfUsar) > 0 Then
                        baseNome = prefixo & "_" & pedidoUsar & "_NF " & nfUsar & "_" & fornecedorUsar
                    Else
                        baseNome = prefixo & "_" & pedidoUsar & "_NF_" & fornecedorUsar
                    End If
                    
                    Dim novoNome As String
                    Select Case tipo
                        Case "XML"
                            ' Mesmo padrão da NF, só muda a extensão
                            novoNome = baseNome & "_XML.xml"    ' ext já é ".xml", manter consistência
                        Case "BOLETO"
                            ' Mesmo padrão da NF + sufixo para não conflitar com o PDF da NF
                            novoNome = baseNome & "_BOLETO" & ext
                        Case "NF"
                            novoNome = baseNome & ext
                        Case Else
                            If LOG_DEBUG Then Debug.Print "Ignorado (tipo inesperado): "; nomeArq
                            GoTo ProximoAnexo
                    End Select
                    
                    Dim destino As String
                    destino = pastaPedido & novoNome
                    
                    ' ==== NOVO: impedir duplicidade ====
                    If SKIP_IF_EXISTS Then
                        If FileExists(destino) Or savedMap.Exists(LCase$(destino)) Then
                            If LOG_DEBUG Then Debug.Print "Ignorado (já existe): "; destino
                            GoTo ProximoAnexo
                        End If
                    End If
                    
                    att.SaveAsFile destino
                    savedMap(LCase$(destino)) = True
                    anexosBaixadosNoGeral = True
                    salvouNesteEmail = True
                    If LOG_DEBUG Then Debug.Print "Salvo: "; destino
ProximoAnexo:
                Next i
                
                ' Marcar como lido somente se ESTE e-mail teve anexo salvo
                If MARK_AS_READ And salvouNesteEmail Then
                    mail.UnRead = False
                    mail.Save
                End If
            End If
        End If
ProximoItem:
    Next item
    
    If anexosBaixadosNoGeral Then
        MsgBox "Anexos baixados e renomeados com sucesso!", vbInformation, "Sucesso"
    Else
        MsgBox "Nenhum anexo elegível (PDF/XML) encontrado em não lidos.", vbInformation, "Aviso"
    End If

Finalizar:
    On Error Resume Next
    Set att = Nothing
    Set mail = Nothing
    Set filteredItems = Nothing
    Set olSubFolder = Nothing
    Set olInbox = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing
    Exit Sub

TratamentoErro:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Falha"
    Resume Finalizar
End Sub

' ============================
' ======= Funções util =======
' ============================

Private Function TryGetSubFolder(ns As Outlook.NameSpace, inbox As Outlook.MAPIFolder, _
                                 topLevel As String, subLevel As String) As Outlook.MAPIFolder
    On Error Resume Next
    Dim fld As Outlook.MAPIFolder
    
    ' 1) No topo do mailbox
    Set fld = ns.Folders(topLevel)
    If Not fld Is Nothing Then
        Set fld = fld.Folders(subLevel)
        If Not fld Is Nothing Then
            Set TryGetSubFolder = fld
            Exit Function
        End If
    End If
    
    ' 2) Dentro da Caixa de Entrada
    Set fld = inbox.Folders(topLevel)
    If Not fld Is Nothing Then
        Set fld = fld.Folders(subLevel)
        If Not fld Is Nothing Then
            Set TryGetSubFolder = fld
            Exit Function
        End If
    End If
    On Error GoTo 0
End Function

Private Function NzStr(ByVal s As Variant) As String
    If IsNull(s) Or IsEmpty(s) Then
        NzStr = ""
    Else
        NzStr = CStr(s)
    End If
End Function

' -------- Normalização --------

Private Function NormalizeForLines(ByVal s As String) As String
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

Private Function RemoveAllWhitespace(ByVal s As String) As String
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

' -------- RegEx helpers --------

Private Function RegexGetFirst(ByVal texto As String, ByVal pattern As String) As String
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
            RegexGetFirst = Trim$(m.SubMatches(0))
        Else
            RegexGetFirst = Trim$(m.Value)
        End If
    Else
        RegexGetFirst = ""
    End If
End Function

Private Function LabelValuePatternTolerant(ByVal label As String) As String
    LabelValuePatternTolerant = label & "\s*[:\-–—]?\s*([^\r\n]+)"
End Function

Private Function LabelValuePatternAfterCompact(ByVal label As String) As String
    LabelValuePatternAfterCompact = label & "[:\-–—]*\s*([0-9\.\-\/]+)"
End Function

Private Function RegexGetLineValue(ByVal texto As String, ByVal label As String) As String
    RegexGetLineValue = RegexGetFirst(texto, LabelValuePatternTolerant(label))
End Function

' -------- Extras --------

Private Function OnlyDigits(ByVal s As String) As String
    Dim i As Long, ch As String, r As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then r = r & ch
    Next i
    OnlyDigits = r
End Function

Private Function EhPedidoOcValido(ByVal codigo As String) As Boolean
    EhPedidoOcValido = (Len(codigo) = 10 And Left$(codigo, 2) = "45")
End Function

Private Function DetectarPrefixo(ByVal corpo As String, ByVal assunto As String) As String
    ' Regra simples: achou "FPP" no assunto/corpo -> "FPP", senão "CG"
    If InStr(1, assunto, "FPP", vbTextCompare) > 0 Or InStr(1, corpo, "FPP", vbTextCompare) > 0 Then
        DetectarPrefixo = "FPP"
    Else
        DetectarPrefixo = "CG"
    End If
End Function

Private Function ObterExtensao(ByVal fileName As String) As String
    Dim p As Long: p = InStrRev(fileName, ".")
    If p > 0 Then
        ObterExtensao = Mid$(fileName, p + 1)
    Else
        ObterExtensao = ""
    End If
End Function

' -------- Checagem específica de PO (Purchase Order) --------
' Cobre: "PO123", "reqPO123", "PO_123", "PO-123", "PO 123",
'        "purchase order", "pedido de compra", "ordem de compra"
Private Function IsArquivoPO(ByVal nmLower As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    With re
        .IgnoreCase = True
        .Global = False
        .Multiline = False
        .pattern = "po\d+|(^|[^A-Za-z0-9])po([^A-Za-z0-9]|$)|purchase[\s_\-]*order|pedido[\s_\-]*de[\s_\-]*compra|ordem[\s_\-]*de[\s_\-]*compra"
    End With
    IsArquivoPO = re.Test(nmLower)
End Function

' -------- Detecção do tipo do anexo --------
Private Function DetectarTipoAnexo(ByVal nomeArq As String, ByVal ext As String) As String
    Dim nm As String: nm = LCase$(nomeArq)
    
    ' XML -> NFe
    If ext = ".xml" Then
        DetectarTipoAnexo = "XML"
        Exit Function
    End If
    
    ' Se o nome indicar PO -> invalidar
    If IsArquivoPO(nm) Then
        DetectarTipoAnexo = "INVALIDO_PO"
        Exit Function
    End If
    
    If ext = ".pdf" Then
        ' BOLETO
        If InStr(nm, "boleto") > 0 Or InStr(nm, "linha digit") > 0 Then
            DetectarTipoAnexo = "BOLETO"
            Exit Function
        End If
        
        ' NF: usa regex pra evitar falso positivo em "info"
        Dim re As Object
        Set re = CreateObject("VBScript.RegExp")
        With re
            .IgnoreCase = True
            .Global = False
            .Multiline = False
            .pattern = "danfe|nfe|nota[\s_\-]*fiscal|(^|[^A-Za-z])nf(\d|[^A-Za-z]|$)"
        End With
        If re.Test(nm) Then
            DetectarTipoAnexo = "NF"
            Exit Function
        End If
        
        ' PDF que não é BOLETO nem NF -> não baixar
        DetectarTipoAnexo = "DESCONHECIDO"
        Exit Function
    End If
    
    ' Outras extensões (se passarem pelo filtro) -> não baixar
    DetectarTipoAnexo = "DESCONHECIDO"
End Function

Private Function SafeFileComponent(ByVal s As String) As String
    Dim t As String, i As Long, ch As Integer
    t = s
    ' Remove caracteres inválidos para nome de arquivo
    t = Replace(t, "\", " ")
    t = Replace(t, "/", " ")
    t = Replace(t, ":", " ")
    t = Replace(t, "*", " ")
    t = Replace(t, "?", " ")
    t = Replace(t, """", " ")
    t = Replace(t, "<", " ")
    t = Replace(t, ">", " ")
    t = Replace(t, "|", " ")
    ' Remove caracteres de controle (ASCII < 32)
    Dim sb As String
    For i = 1 To Len(t)
        ch = Asc(Mid$(t, i, 1))
        If ch >= 32 Then sb = sb & Chr$(ch)
    Next i
    sb = Trim$(sb)
    Do While InStr(sb, "  ") > 0
        sb = Replace(sb, "  ", " ")
    Loop
    SafeFileComponent = sb
End Function

Private Function FileExists(ByVal fullPath As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir$(fullPath, vbNormal)) > 0)
End Function

Private Sub EnsureFolderExistsRecursive(ByVal path As String)
    ' Cria a árvore de pastas se não existir
    Dim fso As Object, parts() As String, cur As String, i As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(path) = 0 Then Exit Sub
    If Right$(path, 1) = "\" Then path = Left$(path, Len(path) - 1)
    parts = Split(path, "\")
    If UBound(parts) < 1 Then Exit Sub
    cur = parts(0)
    For i = 1 To UBound(parts)
        cur = cur & "\" & parts(i)
        If Not fso.FolderExists(cur) Then
            On Error Resume Next
            fso.CreateFolder cur
            On Error GoTo 0
        End If
    Next i
End Sub

Private Function ExtrairNumeroProvavelDeNF(ByVal nomeArquivo As String) As String
    ' Busca sequências de 6–10 dígitos no nome do arquivo (heurística)
    Dim re As Object, mc As Object
    Set re = CreateObject("VBScript.RegExp")
    With re
        .pattern = "(\d{6,10})"
        .Global = True
        .IgnoreCase = True
    End With
    If re.Test(nomeArquivo) Then
        Set mc = re.Execute(nomeArquivo)
        ExtrairNumeroProvavelDeNF = mc(0).SubMatches(0)
    Else
        ExtrairNumeroProvavelDeNF = ""
    End If
End Function

