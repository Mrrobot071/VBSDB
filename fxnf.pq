
let
    Path = "C:\Users\lucas.nathan\OneDrive - ALLOS\Área de Trabalho\Anexostudo",
    KeepOnlyWithPedido = false,
    MinHitsPO = 3,
    KnownTomadorCNPJ = "14.804.587/0001-04",

    Nz = (t as nullable text) as text =>
        if t = null then "" else Text.From(t),

    OnlyDigits = (t as nullable text) as nullable text =>
        if t = null then null else Text.Select(Text.From(t), {"0".."9"}),

    TrimZerosLeft = (t as nullable text) as nullable text =>
        if t = null then null else Text.TrimStart(t, {"0"}),

    CleanWS = (t as nullable text) as text =>
        if t = null then ""
        else Text.Trim(Text.Replace(Text.Replace(Text.Replace(Text.Clean(t), "#(lf)", " "), "#(cr)", " "), "#(tab)", " ")),

    NormalizeNoSpace = (t as nullable text) as text =>
        Text.Upper(Text.Remove(CleanWS(t), {" "})),    

    NormalizeKeepSpace = (t as nullable text) as text =>
        Text.Upper(CleanWS(t)),

    ContainsAny = (hay as nullable text, needles as list) as logical =>
        let
            h = if hay = null then "" else hay
        in
            List.AnyTrue(List.Transform(needles, each Text.Contains(h, _, Comparer.OrdinalIgnoreCase))),

    // =====================================================
    // ✅ PARSER MONETÁRIO ESTRITO (NUNCA "perde" os decimais)
    // Aceita:
    // - "8.687,74" (pt-BR) => 8687.74
    // - "420,00"  (pt-BR) => 420.00
    // - "420.00"  (en-US) => 420.00
    // Rejeita/neutraliza códigos tipo "01.01.01" via filtro no tokenizador.
    // =====================================================
    ParseMoneyStrict = (tok as nullable text) as nullable number =>
        let
            t0 = if tok = null then null else Text.Trim(tok),
            t1 = if t0 = null then null else Text.Select(t0, {"0".."9",".",","}),
            hasComma = if t1 = null then false else Text.Contains(t1, ","),
            hasDot   = if t1 = null then false else Text.Contains(t1, "."),

            // define separador decimal
            decSep =
                if hasComma then ","
                else if hasDot then "."
                else null,

            // separador de milhar provável (oposto do decimal)
            grpSep =
                if decSep = "," then "."
                else if decSep = "." then ","
                else null,

            parts =
                if t1 = null or decSep = null
                then { t1 }
                else Text.Split(t1, decSep),

            intRaw = if List.Count(parts) >= 1 then parts{0} else "",
            decRaw = if List.Count(parts) >= 2 then parts{1} else "",

            // remove separador de milhar do inteiro (ex.: 8.687 -> 8687)
            intNoGrp =
                if grpSep <> null
                then Text.Replace(intRaw, grpSep, "")
                else intRaw,

            intDigits = Text.Select(intNoGrp, {"0".."9"}),
            decDigits = Text.Select(decRaw, {"0".."9"}),

            // força 2 casas (se vier vazio, vira "00"; se vier "7", vira "70")
            dec2 = Text.Start(decDigits & "00", 2),

            composed =
                if intDigits = null or intDigits = "" then null
                else intDigits & "." & dec2,

            n = if composed = null then null else try Number.FromText(composed, "en-US") otherwise null
        in
            n,

    // =====================================================
    // CNPJ helpers
    // =====================================================
    FormatCNPJ = (digits as nullable text) as nullable text =>
        if digits = null or Text.Length(digits) <> 14 then null
        else
            Text.Range(digits, 0, 2) & "." &
            Text.Range(digits, 2, 3) & "." &
            Text.Range(digits, 5, 3) & "/" &
            Text.Range(digits, 8, 4) & "-" &
            Text.Range(digits, 12, 2),

    IsValidCNPJDigits = (digits as nullable text) as logical =>
        let
            d = if digits = null then null else Text.Select(digits, {"0".."9"}),
            okLen = d <> null and Text.Length(d) = 14,
            notRepeated =
                if not okLen then false
                else List.Count(List.Distinct(Text.ToList(d))) > 1,

            nums =
                if not okLen then {}
                else List.Transform(Text.ToList(d), each Number.FromText(_)),

            base12 = if List.Count(nums) = 14 then List.FirstN(nums, 12) else {},
            w1 = {5,4,3,2,9,8,7,6,5,4,3,2},
            s1 = if List.Count(base12) = 12 then List.Sum(List.Transform(List.Positions(base12), each base12{_} * w1{_})) else -1,
            r1 = if s1 < 0 then -1 else Number.Mod(s1, 11),
            dv1 = if r1 < 2 then 0 else 11 - r1,

            base13 = if List.Count(nums) = 14 then List.FirstN(nums, 13) else {},
            w2 = {6,5,4,3,2,9,8,7,6,5,4,3,2},
            s2 = if List.Count(base13) = 13 then List.Sum(List.Transform(List.Positions(base13), each base13{_} * w2{_})) else -1,
            r2 = if s2 < 0 then -1 else Number.Mod(s2, 11),
            dv2 = if r2 < 2 then 0 else 11 - r2,

            valid =
                okLen and
                notRepeated and
                nums{12} = dv1 and
                nums{13} = dv2
        in
            valid,

    // =====================================================
    // Nome/PO/Boleto
    // =====================================================
    IsBoletoByName = (name as nullable text) as logical =>
        if name = null then false else Text.Contains(Text.Lower(name), "boleto"),

    IsPure10DigitBaseName = (name as nullable text) as logical =>
        if name = null then false
        else
            let
                base =
                    if Text.EndsWith(name, ".pdf") or Text.EndsWith(name, ".PDF")
                    then Text.Start(name, Text.Length(name) - 4)
                    else name,
                isOnlyDigits = Text.Length(Text.Remove(base, {"0".."9"})) = 0
            in
                isOnlyDigits and Text.Length(base) = 10,

    IsPOByName = (name as nullable text) as logical =>
        if name = null then false
        else
            let
                nAll = Text.Lower(name),
                nPad = " " & Text.Replace(Text.Replace(Text.Replace(nAll, "_", " "), "-", " "), ".", " ") & " "
            in
                Text.Contains(nAll, "pedido")
                or Text.Contains(nAll, "ordem de compra")
                or Text.Contains(nAll, "ordemdecompra")
                or Text.Contains(nAll, "xml")
                or Text.Contains(nAll, "boleto")
                or Text.Contains(nAll, "(po)")
                or Text.StartsWith(nAll, "po_")
                or Text.EndsWith(nAll, "_po")
                or Text.Contains(nPad, " po ")
                or IsPure10DigitBaseName(name),

    // =====================================================
    // PDF -> texto
    // =====================================================
    TableToText = (tbl as nullable table) as text =>
        if tbl = null then ""
        else
            Text.Combine(
                List.Transform(
                    Table.ToColumns(tbl),
                    each Text.Combine(List.Transform(_, Nz), " ")
                ),
                " "
            ),

    PdfToTextAll = (pdf as any) as text =>
        let
            pages = try Table.Column(pdf, "Data") otherwise {},
            joined = Text.Combine(List.Transform(pages, each TableToText(_)), " ")
        in
            NormalizeKeepSpace(joined),

    PdfToTextPg1_NoSpace = (pdf as any) as text =>
        let
            pages = try Table.Column(pdf, "Data") otherwise {},
            pg1 = if List.Count(pages) > 0 then pages{0} else null,
            txt = TableToText(pg1)
        in
            NormalizeNoSpace(txt),

    PdfToTextPg1_KeepSpace = (pdf as any) as text =>
        let
            pages = try Table.Column(pdf, "Data") otherwise {},
            pg1 = if List.Count(pages) > 0 then pages{0} else null,
            txt = TableToText(pg1)
        in
            NormalizeKeepSpace(txt),

    // =====================================================
    // PO por conteúdo
    // =====================================================
    IsPOByContentStrong = (t as nullable text, optional minHits as nullable number) as logical =>
        let
            s = if t = null then "" else NormalizeNoSpace(t),
            markers = {
                "ORDEMDECOMPRA(PO)","ORDEMDECOMPRA","NºPO","NPO","NUMERODAPO",
                "DADOSDOFORNECEDOR","DADOSPARAFATURAMENTO","DADOSPARACOBRANÇA","DADOSPARACOBRANCA",
                "ENDEREÇODEENTREGA","ENDERECODEENTREGA",
                "AUTORIZAMOSOFORNECIMENTODOSITENSABAIXO",
                "CONDICOESDEPAGAMENTO","CONDIÇÕESDEPAGAMENTO",
                "DADOSDEPROCESSO","CALENDARIODEPAGAMENTOS","CALENDÁRIODEPAGAMENTOS",
                "TRANSPORTADORA","ESTAORDEMDECOMPRAFOIAPROVADAELETRONICAMENTE"
            },
            hits = List.Sum(List.Transform(markers, each if Text.Contains(s, _) then 1 else 0)),
            need = if minHits = null then 3 else Number.From(minHits)
        in
            hits >= need,

    // =====================================================
    // Tokens / extrações de nome
    // =====================================================
    FirstNumberToken = (s as text) as nullable text =>
        let
            L = Text.Length(s),
            idxDigits = List.Select({0..L-1}, each List.Contains({"0".."9"}, Text.Range(s, _, 1))),
            start = if List.Count(idxDigits) > 0 then idxDigits{0} else -1,
            chars =
                if start < 0 then {}
                else
                    List.Generate(
                        () => start,
                        each _ < L and List.Contains({"0".."9",".","-","/"}, Text.Range(s, _, 1)),
                        each _ + 1,
                        each Text.Range(s, _, 1)
                    ),
            token = if List.Count(chars) > 0 then Text.Combine(chars, "") else null
        in
            token,

    PedidoFromName = (name as text) as nullable text =>
        let
            parts = Text.Split(name, "_"),
            p = List.First(List.Select(parts, each Text.Length(OnlyDigits(_)) = 10), null)
        in
            OnlyDigits(p),

    ExtractPedidoFromText = (t as nullable text) as nullable text =>
        let
            U = if t = null then "" else NormalizeKeepSpace(t),
            markers = {
                "PEDIDO DE COMPRA",
                "PEDIDO",
                "ORDEM DE COMPRA",
                "NUMERO DO PEDIDO",
                "NÚMERO DO PEDIDO",
                "Nº DE PROTOCOLO",
                "NO DE PROTOCOLO",
                "PROTOCOLO"
            },

            tryMarker = (m as text) as nullable text =>
                let
                    pos = Text.PositionOf(U, m, Occurrence.First)
                in
                    if pos < 0 then null
                    else
                        let
                            segRaw = Text.Middle(U, pos + Text.Length(m), 80),
                            segClean = Text.Replace(Text.Replace(Text.Replace(segRaw, ":", " "), "=", " "), "#", " "),
                            digits = Text.Select(segClean, {"0".."9"}),
                            cand =
                                if Text.Length(digits) >= 10
                                then Text.Start(digits, 10)
                                else null
                        in
                            cand,

            cands = List.RemoveNulls(List.Transform(markers, each tryMarker(_))),
            result = if List.Count(cands) > 0 then cands{0} else null
        in
            result,

    NFFromName = (name as text) as nullable text =>
        let
            base = if Text.EndsWith(name, ".pdf") or Text.EndsWith(name, ".PDF")
                   then Text.Start(name, Text.Length(name)-4)
                   else name,
            u = Text.Upper(base),
            markers = {"NFSE","NFS-E","NFE","NF-E","NF","DANFE","DANFSE"},

            GetDigitsAfter = (tail as text) as nullable text =>
                let
                    tok = FirstNumberToken(tail),
                    onlyDigits = if tok = null then null else Text.Select(tok, {"0".."9"}),
                    cleaned = if onlyDigits = null then null else TrimZerosLeft(onlyDigits)
                in
                    if cleaned <> null and Text.Length(cleaned) >= 2 then cleaned else null,

            tryByMarkers =
                List.RemoveNulls(
                    List.Transform(
                        markers,
                        (m) =>
                            let
                                pos = Text.PositionOf(u, m, Occurrence.First)
                            in
                                if pos < 0 then null
                                else
                                    let tail = Text.Middle(base, pos + Text.Length(m))
                                    in GetDigitsAfter(tail)
                    )
                ),
            byMarker = if List.Count(tryByMarkers) > 0 then tryByMarkers{0} else null,

            digitsAndSpace =
                Text.Combine(
                    List.Transform(Text.ToList(base), each if List.Contains({"0".."9"}, _) then _ else " "),
                    ""
                ),
            parts = List.Select(Text.Split(digitsAndSpace, " "), each Text.Length(_) > 0),
            pedidos10 = List.Select(parts, each Text.Length(_) = 10),
            candRaw = List.Select(parts, each let L = Text.Length(_) in L >= 2 and L <= 9),
            candNoPO = List.Select(candRaw, each not List.Contains(pedidos10, _)),
            picked = if List.Count(candNoPO) > 0 then candNoPO{List.Count(candNoPO)-1} else null,

            result = if byMarker <> null then byMarker else picked
        in
            result,

    FornecedorFromName = (name as text) as nullable text =>
        let
            base = if Text.EndsWith(name, ".pdf") or Text.EndsWith(name, ".PDF")
                   then Text.Start(name, Text.Length(name)-4)
                   else name,

            u = Text.Upper(base),
            markers = {"NFSE","NFS-E","NFE","NF-E","NF","DANFE","DANFSE"},

            posList =
                List.RemoveNulls(
                    List.Transform(
                        markers,
                        each
                            let p = Text.PositionOf(u, _, Occurrence.First)
                            in if p >= 0 then [m=_, p=p] else null
                    )
                ),
            posMin = if List.Count(posList) = 0 then null else List.Min(List.Transform(posList, each _[p])),
            marker = if posMin = null then null else List.First(List.Select(posList, each _[p] = posMin))[m],

            afterMarker =
                if marker = null then null
                else Text.Middle(base, posMin + Text.Length(marker)),

            tail0 = if afterMarker = null then null else Text.TrimStart(afterMarker, {" ","_","-",".",":","(",")"}),
            nfTok = if tail0 = null then null else FirstNumberToken(tail0),
            tail1 =
                if tail0 = null then null
                else if nfTok = null then tail0
                else Text.TrimStart(Text.Range(tail0, Text.Length(nfTok)), {" ","_","-",".",":","(",")"}),

            fornecedorRaw = if tail1 = null then null else Text.Trim(tail1),
            fornecedor = if fornecedorRaw = null or fornecedorRaw = "" then null else fornecedorRaw
        in
            fornecedor,

    DetectTipo = (name as nullable text, txtPg1 as nullable text, txtAll as nullable text) as text =>
        let
            N = NormalizeKeepSpace(name),
            U1 = NormalizeKeepSpace(txtPg1),
            UAll = NormalizeKeepSpace(txtAll),
            Combined = N & " " & U1 & " " & UAll,
            HasXML = ContainsAny(Combined, {"XML"}),
            HasProduto = ContainsAny(Combined, {"PRODUTO","PRODUTOS"})
        in
            if HasXML or HasProduto then "PRODUTO" else "SERVIÇO",

    // =====================================================
    // ✅ TOKENIZAÇÃO MONETÁRIA CORRIGIDA
    // - só aceita tokens com:
    //   a) vírgula (pt-BR), ou
    //   b) UM único ponto como decimal (en-US)
    // - rejeita "01.01.01" (múltiplos pontos)
    // =====================================================
    MoneyTokensFromText = (seg as nullable text) as list =>
        let
            s = if seg = null then "" else seg,
            mapped =
                Text.Combine(
                    List.Transform(
                        Text.ToList(s),
                        each if List.Contains({"0".."9",".",","}, _) then _ else "|"
                    ),
                    ""
                ),
            toks0 = List.Select(Text.Split(mapped, "|"), each Text.Length(_) > 0),

            IsOkMoneyToken = (tok as text) as logical =>
                let
                    d = Text.Select(tok, {"0".."9"}),
                    hasComma = Text.Contains(tok, ","),
                    dotCount = List.Count(Text.Split(tok, ".")) - 1,
                    afterLastDot = if Text.Contains(tok, ".") then Text.AfterDelimiter(tok, ".", {0, RelativePosition.FromEnd}) else "",
                    looksDecimalDot = dotCount = 1 and (Text.Length(afterLastDot) = 1 or Text.Length(afterLastDot) = 2),
                    ok = Text.Length(d) >= 3 and (hasComma or looksDecimalDot)
                in
                    ok,

            toks1 = List.Select(toks0, each IsOkMoneyToken(_)),

            recs =
                List.RemoveNulls(
                    List.Transform(
                        toks1,
                        each
                            let
                                n = ParseMoneyStrict(_)
                            in
                                if n <> null and n > 0 then [Tok=_, Num=n] else null
                    )
                )
        in
            recs,

    ExtractValor = (t as nullable text) as nullable number =>
        let
            U0 = if t = null then null else NormalizeKeepSpace(t),
            U  = if U0 = null then "" else U0,

            anchors = {
                "VALOR A PAGAR","VALOR DA NOTA FISCAL","VALOR DA NOTA",
                "VALOR TOTAL DA NFS-E","VALOR TOTAL DA NFSE","VALOR TOTAL DA NFS E",
                "VALOR TOTAL DO DOCUMENTO","VALOR TOTAL DA NOTA","TOTAL DA NOTA","V. TOTAL DA NOTA",
                "TOTAL DO DOCUMENTO","VALOR TOTAL","VALOR TOTAL:",
                "VALOR DO SERVIÇO","VALOR DOS SERVIÇOS","VALOR DO SERVICO","VALOR DOS SERVICOS",
                "VALOR TOTAL DOS SERVIÇOS","VALOR TOTAL DOS SERVICOS",
                "TOTAL A PAGAR","TOTAL A RECEBER","TOTAL GERAL"
            },

            negWords = {
                "RETEN","ISS","IRRF","INSS","PIS","COFINS","CSLL","SEST","SENAT","IBF","IBPT","ICMS",
                "BASE DE CÁLCULO","BASE DE CALCULO","ALIQUOTA","ALÍQUOTA",
                "DESCONTO","DESCONTOS","DEDUÇÃO","DEDUCAO"
            },

            ScoreWindow = (seg as text) as number =>
                let
                    negHits = List.Sum(List.Transform(negWords, each if Text.Contains(seg, _, Comparer.OrdinalIgnoreCase) then 1 else 0))
                in
                    100 - (negHits * 25),

            CandFromAnchor = (a as text) as nullable record =>
                let
                    p = Text.PositionOf(U, a, Occurrence.First)
                in
                    if p < 0 then null
                    else
                        let
                            tail = Text.Middle(U, p + Text.Length(a), 200),
                            pRS = Text.PositionOf(tail, "R$", Occurrence.First),
                            afterRS = if pRS >= 0 then Text.Middle(tail, pRS + 2, 80) else tail,

                            toksAfterRS = MoneyTokensFromText(afterRS),
                            n1 = if List.Count(toksAfterRS) > 0 then toksAfterRS{0}[Num] else null,

                            toksTail = if n1 <> null then {} else MoneyTokensFromText(tail),
                            n2 = if n1 <> null then n1 else if List.Count(toksTail) > 0 then toksTail{0}[Num] else null,

                            ctxStart = if p - 40 < 0 then 0 else p - 40,
                            ctxLen   = if (ctxStart + Text.Length(a) + 260) > Text.Length(U) then Text.Length(U) - ctxStart else Text.Length(a) + 260,
                            ctx      = Text.Range(U, ctxStart, ctxLen),
                            sc       = ScoreWindow(ctx)
                        in
                            if n2 = null then null else [Val=n2, Score=sc, Pos=p],

            candByAnch = List.RemoveNulls(List.Transform(anchors, each CandFromAnchor(_))),

            chosenAnch =
                if List.Count(candByAnch) = 0 then null
                else
                    let
                        tbl    = Table.FromRecords(candByAnch),
                        sorted = Table.Sort(tbl, {{"Score", Order.Descending}, {"Pos", Order.Ascending}})
                    in
                        sorted{0}[Val],

            // fallback: pega o MAIOR valor monetário do texto (já filtrado corretamente)
            allMoney = MoneyTokensFromText(U),
            valsAll  = List.Transform(allMoney, each [Num]),
            maxAll   = if List.Count(valsAll) > 0 then List.Max(valsAll) else null,

            result = if chosenAnch <> null then chosenAnch else maxAll
        in
            result,

    ExtractDatesBR = (t as text) as list =>
        let
            U = t,
            L = Text.Length(U),
            poss = if L < 10 then {} else {0..(L-10)},
            ds =
                List.RemoveNulls(
                    List.Transform(poss, (i) =>
                        let
                            s = Text.Range(U, i, 10),
                            ok = Text.Range(s, 2, 1) = "/" and Text.Range(s, 5, 1) = "/",
                            dTxt = Text.Range(s, 0, 2),
                            mTxt = Text.Range(s, 3, 2),
                            yTxt = Text.Range(s, 6, 4),
                            d = try Number.FromText(dTxt) otherwise -1,
                            m = try Number.FromText(mTxt) otherwise -1,
                            y = try Number.FromText(yTxt) otherwise -1
                        in
                            if ok and d >= 1 and d <= 31 and m >= 1 and m <= 12 and y >= 2000 and y <= 2100
                            then try #date(y, m, d) otherwise null
                            else null
                    )
                )
        in
            ds,

    ExtractEmissao = (t as text) as nullable date =>
        let
            U = NormalizeKeepSpace(t),
            anchors = {"DATA E HORA DA EMISSÃO","DATA DA EMISSÃO","DATA DE EMISSÃO","EMISSÃO","EMISSAO","COMPETÊNCIA DA NFS-E","COMPETÊNCIA"},
            byAnch =
                List.RemoveNulls(
                    List.Transform(anchors, (a) =>
                        let pos = Text.PositionOf(U, a, Occurrence.First)
                        in
                            if pos < 0 then null
                            else
                                let seg = Text.Middle(U, pos, 140),
                                    ds  = ExtractDatesBR(seg)
                                in if List.Count(ds) > 0 then ds{0} else null
                    )
                ),
            d1 = if List.Count(byAnch) > 0 then byAnch{0} else null,
            allDs = ExtractDatesBR(U),
            d2 = if d1 <> null then d1 else if List.Count(allDs) > 0 then allDs{0} else null
        in
            d2,

    EmissaoFromName = (name as text) as nullable date =>
        let
            rep = Text.Replace(Text.Replace(Text.Replace(name, "-", "/"), "_", "/"), ".", "/"),
            ds = ExtractDatesBR(rep),
            ok = List.Select(ds, each Date.Year(_) >= 2024)
        in
            if List.Count(ok) > 0 then ok{0} else null,

    ExtractCNPJ = (t as text) as nullable text =>
        let
            U = NormalizeKeepSpace(t),
            L = Text.Length(U),
            KnownDigits = OnlyDigits(KnownTomadorCNPJ),

            fmtCandidates =
                if L < 18 then {}
                else
                    List.RemoveNulls(
                        List.Transform(
                            {0..(L-18)},
                            (i) =>
                                let
                                    s18 = Text.Middle(U, i, 18),
                                    digits = OnlyDigits(s18),
                                    isFmt =
                                        Text.Middle(s18, 2, 1) = "." and
                                        Text.Middle(s18, 6, 1) = "." and
                                        Text.Middle(s18, 10, 1) = "/" and
                                        Text.Middle(s18, 15, 1) = "-" and
                                        digits <> null and Text.Length(digits) = 14
                                in if isFmt then digits else null
                        )
                    ),

            rawCandidates =
                if L < 14 then {}
                else
                    List.RemoveNulls(
                        List.Transform(
                            {0..(L-14)},
                            (i) =>
                                let
                                    s14 = Text.Middle(U, i, 14),
                                    digits = OnlyDigits(s14)
                                in if digits <> null and Text.Length(digits) = 14 then digits else null
                        )
                    ),

            allCandidates0 = fmtCandidates & rawCandidates,
            allCandidates1 = List.Select(allCandidates0, each _ <> null and _ <> KnownDigits and List.Count(List.Distinct(Text.ToList(_))) > 1),
            uniqDigits = List.Distinct(allCandidates1),
            validDigits = List.Select(uniqDigits, each IsValidCNPJDigits(_)),
            picked = if List.Count(validDigits) > 0 then validDigits{0} else if List.Count(uniqDigits) > 0 then uniqDigits{0} else null
        in
            FormatCNPJ(picked),

    ExtractNumServico = (t as text) as nullable text =>
        let
            U = NormalizeKeepSpace(t),
            markers = {"NÚMERO DE SERVIÇO","NUMERO DE SERVIÇO","NUMERO DE SERVICO","NRO SERVIÇO","NRO SERVICO","NO SERVIÇO","NO SERVICO"},
            cands =
                List.RemoveNulls(
                    List.Transform(markers, (m) =>
                        let
                            pos = Text.PositionOf(U, m, Occurrence.First)
                        in
                            if pos < 0 then null
                            else
                                let seg = Text.Middle(U, pos, 120),
                                    nums = Text.Select(seg, {"0".."9"})
                                in if Text.Length(nums) >= 3 then TrimZerosLeft(nums) else null
                    )
                )
        in
            if List.Count(cands) > 0 then cands{0} else null,

    // =========================
    // PIPELINE
    // =========================
    Source = Folder.Files(Path),
    OnlyPDF = Table.SelectRows(Source, each Text.Lower([Extension]) = ".pdf"),

    AddIsPO_Name = Table.AddColumn(OnlyPDF, "IsPO_Name", each IsPOByName([Name]), type logical),
    NoPO_Name = Table.SelectRows(AddIsPO_Name, each [IsPO_Name] = false),

    AddPdf = Table.AddColumn(NoPO_Name, "PdfTbl", each try Pdf.Tables([Content]) otherwise null),

    AddTextAll = Table.AddColumn(AddPdf, "TextAll", each try PdfToTextAll([PdfTbl]) otherwise "", type text),
    AddTextPg1KS = Table.AddColumn(AddTextAll, "TextPg1", each try PdfToTextPg1_KeepSpace([PdfTbl]) otherwise "", type text),
    AddTextPg1NS = Table.AddColumn(AddTextPg1KS, "TextPg1_NoSpace", each try PdfToTextPg1_NoSpace([PdfTbl]) otherwise "", type text),

    AddIgnore = Table.AddColumn(
        AddTextPg1NS,
        "Ignore",
        each IsBoletoByName([Name]) or IsPOByContentStrong([TextAll], MinHitsPO),
        type logical
    ),
    Kept = Table.SelectRows(AddIgnore, each [Ignore] = false),

    AddPedido_Name = Table.AddColumn(Kept, "PEDIDO_NOME", each PedidoFromName([Name]), type text),
    AddPedido_Text = Table.AddColumn(AddPedido_Name, "PEDIDO_TEXTO", each ExtractPedidoFromText([TextAll]), type text),
    AddPedido = Table.AddColumn(
        AddPedido_Text,
        "PEDIDO",
        each if [PEDIDO_NOME] <> null then [PEDIDO_NOME] else [PEDIDO_TEXTO],
        type text
    ),
    RemPedidoAux = Table.RemoveColumns(AddPedido, {"PEDIDO_NOME", "PEDIDO_TEXTO"}),

    AddNF = Table.AddColumn(RemPedidoAux, "NUM_NF", each NFFromName([Name]), type text),

    AddCNPJ = Table.AddColumn(AddNF, "CNPJ", each ExtractCNPJ([TextAll]), type text),

    AddForn = Table.AddColumn(AddCNPJ, "FORNECEDOR", each FornecedorFromName([Name]), type text),

    AddEmiss_Txt = Table.AddColumn(AddForn, "EMISSAO_TXT", each ExtractEmissao([TextAll]), type date),
    AddEmiss = Table.AddColumn(AddEmiss_Txt, "EMISSAO", each if [EMISSAO_TXT] <> null then [EMISSAO_TXT] else EmissaoFromName([Name]), type date),
    RemEmiss_Txt = Table.RemoveColumns(AddEmiss, {"EMISSAO_TXT"}),

    AddTipo = Table.AddColumn(RemEmiss_Txt, "TIPO", each DetectTipo([Name], [TextPg1], [TextAll]), type text),

    // =====================================================
    // ✅ VALOR CORRETO:
    // - VALOR_NUM = número com 2 casas (8687.74)
    // - VALOR     = texto SEMPRE "8.687,74"
    // =====================================================
    AddValorNum = Table.AddColumn(
        AddTipo,
        "VALOR_NUM",
        each
            let
                v1 = ExtractValor([TextPg1]),
                v2 = if v1 = null then ExtractValor([TextAll]) else v1
            in
                if v2 = null then null else Number.Round(v2, 2),
        type number
    ),

    AddValor = Table.AddColumn(
        AddValorNum,
        "VALOR",
        each if [VALOR_NUM] = null then null else Number.ToText([VALOR_NUM], "#,##0.00", "pt-BR"),
        type text
    ),

    AddNServ = Table.AddColumn(AddValor, "NUM_SERVICO", each ExtractNumServico([TextPg1]), type text),

    Selected =
        Table.SelectColumns(
            AddNServ,
            {"Name","PEDIDO","TIPO","NUM_NF","FORNECEDOR","CNPJ","EMISSAO","VALOR","VALOR_NUM","NUM_SERVICO","TextPg1_NoSpace"}
        ),

    FinalRenamed = Table.RenameColumns(Selected, {{"TextPg1_NoSpace","TEXTO_PG1_SEM_ESPACO"}}),

    Final =
        if KeepOnlyWithPedido
        then Table.SelectRows(FinalRenamed, each [PEDIDO] <> null)
        else FinalRenamed,

    #"Linhas Filtradas" = Table.SelectRows(Final, each [PEDIDO] <> null),
    #"Colunas Removidas" = Table.RemoveColumns(#"Linhas Filtradas", {"NUM_SERVICO"}),

    #"Tipo Alterado" = Table.TransformColumnTypes(
        #"Colunas Removidas",
        {{"EMISSAO", type date}, {"VALOR", type text}, {"VALOR_NUM", type number}}
    )
in
    #"Tipo Alterado"
