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

    PtBR_ToNumber = (s as nullable text) as nullable number =>
        let
            raw      = if s = null then null else Text.Trim(s),
            noMilhar = if raw = null then null else Text.Replace(raw, ".", ""),
            dotDec   = if noMilhar = null then null else Text.Replace(noMilhar, ",", "."),
            n        = if dotDec = null then null else try Number.FromText(dotDec) otherwise null
        in
            n,

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

    IsBoletoByName = (name as nullable text) as logical =>
        if name = null then false
        else
            let
                n = Text.Lower(name)
            in
                Text.Contains(n, "boleto"),

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

    ExtractByMarker = (t as text, marker as text) as nullable text =>
        let
            pos = Text.PositionOf(t, marker, Occurrence.First),
            res =
                if pos < 0 then null
                else
                    let
                        segRaw = Text.Middle(t, pos + Text.Length(marker), 120),
                        removeTs = {"NO.","Nº","N°","N.","NUM.","NÚM.","NUMERO","NÚMERO","=" , ":"},
                        segClean = List.Accumulate(removeTs, segRaw, (state, tok) => Text.Replace(state, tok, "")),
                        pos1 = Text.PositionOf(segClean, "SÉRIE", Occurrence.First),
                        pos2 = Text.PositionOf(segClean, "SERIE", Occurrence.First),
                        posList = List.Select({pos1, pos2}, each _ >= 0),
                        cutPos = if List.Count(posList) > 0 then List.Min(posList) else null,
                        seg = if cutPos <> null then Text.Start(segClean, cutPos) else segClean,
                        token = FirstNumberToken(seg),
                        onlyD = if token = null then null else Text.Select(token, {"0".."9"}),
                        cleaned = if onlyD = null or Text.Length(onlyD) < 2 then null else TrimZerosLeft(onlyD)
                    in
                        cleaned
        in
            res,

    DetectTipo = (name as nullable text, txtPg1 as nullable text, txtAll as nullable text) as text =>
        let
            N = NormalizeKeepSpace(name),
            U1 = NormalizeKeepSpace(txtPg1),
            UAll = NormalizeKeepSpace(txtAll),
            Combined = N & " " & U1 & " " & UAll,
            HasXML = ContainsAny(Combined, {"XML"}),
            HasProduto = ContainsAny(Combined, {"PRODUTO","PRODUTOS"}),
            Result = if HasXML or HasProduto then "PRODUTO" else "SERVIÇO"
        in
            Result,

    // =========================
    // CORREÇÃO: EXTRAÇÃO DE VALOR ROBUSTA
    // =========================

    // Quebra um trecho em tokens contíguos contendo só dígitos/. /,
    // evitando concatenar outros números depois do valor monetário.
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
            toks1 =
                List.Select(
                    toks0,
                    each
                        let
                            d = Text.Select(_, {"0".."9"})
                        in
                            Text.Length(d) >= 3
                            and (Text.Contains(_, ",") or Text.Contains(_, "."))
                            and not Text.StartsWith(_, ",")
                            and not Text.EndsWith(_, ",")
                            and not Text.StartsWith(_, ".")
                            and not Text.EndsWith(_, ".")
                ),
            recs =
                List.RemoveNulls(
                    List.Transform(
                        toks1,
                        each
                            let
                                n = PtBR_ToNumber(_),
                                decPart =
                                    if Text.Contains(_, ",")
                                    then Text.AfterDelimiter(_, ",", {0, RelativePosition.FromEnd})
                                    else null
                            in
                                if n <> null
                                   and n > 0
                                   and Text.Length(Text.Select(_, {"0".."9"})) <= 15
                                   and (decPart = null or Text.Length(decPart) <= 2)
                                then [Tok=_, Num=n]
                                else null
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
                "VALOR TOTAL DA NF-E","VALOR TOTAL DA NFE","VALOR TOTAL DA NF E",
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
                    negHits = List.Sum(List.Transform(negWords, each if Text.Contains(seg, _, Comparer.OrdinalIgnoreCase) then 1 else 0)),
                    score   = 100 - (negHits * 25)
                in
                    score,

            CandFromAnchor = (a as text) as nullable record =>
                let
                    p = Text.PositionOf(U, a, Occurrence.First)
                in
                    if p < 0 then null
                    else
                        let
                            tail = Text.Middle(U, p + Text.Length(a), 180),

                            pRS = Text.PositionOf(tail, "R$", Occurrence.First),
                            afterRS = if pRS >= 0 then Text.Middle(tail, pRS + 2, 50) else tail,

                            toksAfterRS = MoneyTokensFromText(afterRS),
                            n1 = if List.Count(toksAfterRS) > 0 then toksAfterRS{0}[Num] else null,

                            toksTail = if n1 <> null then {} else MoneyTokensFromText(tail),
                            n2 = if n1 <> null then n1 else if List.Count(toksTail) > 0 then toksTail{0}[Num] else null,

                            ctxStart = if p - 40 < 0 then 0 else p - 40,
                            ctxLen   = if (ctxStart + Text.Length(a) + 240) > Text.Length(U) then Text.Length(U) - ctxStart else Text.Length(a) + 240,
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

            // fallback 1: primeiro token monetário após cada ocorrência de R$
            partsRS = List.Skip(Text.Split(U, "R$"), 1),
            valsRS =
                List.RemoveNulls(
                    List.Transform(
                        partsRS,
                        each
                            let
                                toks = MoneyTokensFromText(Text.Start(_, 50))
                            in
                                if List.Count(toks) > 0 then toks{0}[Num] else null
                    )
                ),
            lastRS  = if List.Count(valsRS) > 0 then valsRS{List.Count(valsRS)-1} else null,
            maxRS   = if List.Count(valsRS) > 0 then List.Max(valsRS) else null,

            // fallback 2: qualquer token monetário do texto
            allMoney = MoneyTokensFromText(U),
            valsAll  = List.Transform(allMoney, each [Num]),
            maxAll   = if List.Count(valsAll) > 0 then List.Max(valsAll) else null,

            result =
                if chosenAnch <> null then chosenAnch
                else if lastRS <> null then lastRS
                else if maxRS <> null then maxRS
                else maxAll
        in
            result,

    ExtractNumServico = (t as text) as nullable text =>
        let
            U = NormalizeKeepSpace(t),
            markers = {
                "NÚMERO DE SERVIÇO","NUMERO DE SERVIÇO","NUMERO DE SERVICO",
                "NRO SERVIÇO","NRO SERVICO","NO SERVIÇO","NO SERVICO"
            },
            cands =
                List.RemoveNulls(
                    List.Transform(markers, (m) =>
                        let
                            pos = Text.PositionOf(U, m, Occurrence.First)
                        in
                            if pos < 0 then null
                            else
                                let
                                    seg = Text.Middle(U, pos, 120),
                                    nums = Text.Select(seg, {"0".."9"})
                                in
                                    if Text.Length(nums) >= 3 then TrimZerosLeft(nums) else null
                    )
                )
        in
            if List.Count(cands) > 0 then cands{0} else null,

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
            anchors = {
                "DATA E HORA DA EMISSÃO","DATA DA EMISSÃO","DATA DE EMISSÃO","EMISSÃO","EMISSAO","EMISSÃO NF-E","EMISSÃO NFS-E"
            },
            byAnch =
                List.RemoveNulls(
                    List.Transform(anchors, (a) =>
                        let
                            pos = Text.PositionOf(U, a, Occurrence.First)
                        in
                            if pos < 0 then null
                            else
                                let
                                    seg = Text.Middle(U, pos, 120),
                                    ds = ExtractDatesBR(seg)
                                in
                                    if List.Count(ds) > 0 then ds{0} else null
                    )
                ),
            d1 = if List.Count(byAnch) > 0 then byAnch{0} else null,
            allDs = ExtractDatesBR(U),
            ds2024 = List.Select(allDs, each Date.Year(_) >= 2024),
            d2 = if d1 <> null then d1 else if List.Count(ds2024) > 0 then ds2024{0} else null
        in
            d2,

    EmissaoFromName = (name as text) as nullable date =>
        let
            rep = Text.Replace(Text.Replace(Text.Replace(name, "-", "/"), "_", "/"), ".", "/"),
            ds = ExtractDatesBR(rep),
            ok = List.Select(ds, each Date.Year(_) >= 2024)
        in
            if List.Count(ok) > 0 then ok{0} else null,

    ExtractNF = (t as text) as nullable text =>
        let
            p1 = {
                "NÚMERO DA NOTA","NUMERO DA NOTA",
                "NÚMERO DA NFS-E","NUMERO DA NFS-E","NÚMERO DA NFSE","NUMERO DA NFSE","NÚMERO DA NFS E","NUMERO DA NFS E",
                "NFS-E Nº","NFS-E N°","NFS-E NO.","NFS E Nº","NFS E NO.","NFSE Nº","NFSE NO.",
                "NF-E Nº","NF-E N°","NF-E NO.","NFE Nº","NFE N°","NFE NO."
            },
            p2 = {
                "DANFE","NFE","NF-E","NFS-E","NFSE","DANFSE","DOCUMENTO AUXILIAR DA NFS-E","DOCUMENTO AUXILIAR DA NFSE"
            },

            tryP1 = List.RemoveNulls(List.Transform(p1, each ExtractByMarker(t, _))),
            ansP1 = if List.Count(tryP1) > 0 then tryP1{0} else null,

            tryP2 = if ansP1 <> null then {} else List.RemoveNulls(List.Transform(p2, each ExtractByMarker(t, _))),
            ansP2 = if List.Count(tryP2) > 0 then tryP2{0} else null,

            result = if ansP1 <> null then ansP1 else ansP2
        in
            result,

    // =========================
    // EXTRAÇÃO GENÉRICA DE CNPJ
    // =========================
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
                                        digits <> null and
                                        Text.Length(digits) = 14
                                in
                                    if isFmt then [Pos=i, Digits=digits] else null
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
                                in
                                    if digits <> null and Text.Length(digits) = 14 then [Pos=i, Digits=digits] else null
                        )
                    ),

            allCandidates0 = fmtCandidates & rawCandidates,

            allCandidates1 =
                List.Select(
                    allCandidates0,
                    each [Digits] <> null and [Digits] <> KnownDigits
                ),

            allCandidates2 =
                List.Select(
                    allCandidates1,
                    each List.Count(List.Distinct(Text.ToList([Digits]))) > 1
                ),

            uniqDigits = List.Distinct(List.Transform(allCandidates2, each [Digits])),
            uniqCandidates =
                List.Transform(
                    uniqDigits,
                    each List.First(List.Select(allCandidates2, (r) => r[Digits] = _))
                ),

            validCandidates = List.Select(uniqCandidates, each IsValidCNPJDigits([Digits])),
            candidatesToUse = if List.Count(validCandidates) > 0 then validCandidates else uniqCandidates,

            ScoreCandidate = (rec as record) as number =>
                let
                    p = rec[Pos],
                    start = if p - 120 < 0 then 0 else p - 120,
                    lenCtx = if start + 260 > L then L - start else 260,
                    ctx = Text.Range(U, start, lenCtx),

                    posHits = List.Sum(
                        List.Transform(
                            {
                                "EMITENTE","EMISSOR","PRESTADOR","FORNECEDOR",
                                "DADOS DO FORNECEDOR","DADOSDOFORNECEDOR",
                                "CNPJ","CNPJ/CPF","RAZÃO SOCIAL","RAZAO SOCIAL"
                            },
                            each if Text.Contains(ctx, _, Comparer.OrdinalIgnoreCase) then 1 else 0
                        )
                    ),

                    negHits = List.Sum(
                        List.Transform(
                            {
                                "TOMADOR","DESTINATÁRIO","DESTINATARIO",
                                "CLIENTE","PAGADOR","SACADO"
                            },
                            each if Text.Contains(ctx, _, Comparer.OrdinalIgnoreCase) then 1 else 0
                        )
                    ),

                    bonusStrong =
                        if Text.Contains(ctx, "EMITENTE", Comparer.OrdinalIgnoreCase)
                            or Text.Contains(ctx, "EMISSOR", Comparer.OrdinalIgnoreCase)
                            or Text.Contains(ctx, "PRESTADOR", Comparer.OrdinalIgnoreCase)
                            or Text.Contains(ctx, "FORNECEDOR", Comparer.OrdinalIgnoreCase)
                        then 10 else 0,

                    score = (posHits * 10) + bonusStrong - (negHits * 12)
                in
                    score,

            ranked =
                List.Transform(
                    candidatesToUse,
                    each [Pos = [Pos], Digits = [Digits], Score = ScoreCandidate(_)]
                ),

            rankedTable =
                if List.Count(ranked) = 0
                then #table({"Pos","Digits","Score"}, {})
                else Table.FromRecords(ranked),

            sortedTable = Table.Sort(rankedTable, {{"Score", Order.Descending}, {"Pos", Order.Ascending}}),
            bestDigits = if Table.RowCount(sortedTable) > 0 then sortedTable{0}[Digits] else null,
            result = FormatCNPJ(bestDigits)
        in
            result,

    CNPJCandidatesLoose = (t as nullable text) as list =>
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
                                        digits <> null and
                                        Text.Length(digits) = 14
                                in
                                    if isFmt then digits else null
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
                                in
                                    if digits <> null and Text.Length(digits) = 14 then digits else null
                        )
                    ),

            allCandidates =
                List.Distinct(
                    List.Select(
                        fmtCandidates & rawCandidates,
                        each _ <> null and _ <> KnownDigits and List.Count(List.Distinct(Text.ToList(_))) > 1
                    )
                )
        in
            allCandidates,

    ExtractCNPJLoose = (t as nullable text) as nullable text =>
        let
            cands = CNPJCandidatesLoose(t),
            valid = List.Select(cands, each IsValidCNPJDigits(_)),
            picked =
                if List.Count(valid) > 0 then valid{0}
                else if List.Count(cands) > 0 then cands{0}
                else null
        in
            FormatCNPJ(picked),

    // =========================
    // CNPJ DO BLOCO CERTO
    // =========================
    FindFirstMarkerRecord = (txt as text, markers as list) as nullable record =>
        let
            found =
                List.RemoveNulls(
                    List.Transform(
                        markers,
                        each
                            let p = Text.PositionOf(txt, _, Occurrence.First)
                            in if p >= 0 then [Marker=_, Pos=p] else null
                    )
                ),
            minPos = if List.Count(found) = 0 then null else List.Min(List.Transform(found, each [Pos])),
            rec = if minPos = null then null else List.First(List.Select(found, each [Pos] = minPos))
        in
            rec,

    FirstValidCNPJInText = (t as nullable text) as nullable text =>
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
                                        digits <> null and
                                        Text.Length(digits) = 14
                                in
                                    if isFmt then digits else null
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
                                in
                                    if digits <> null and Text.Length(digits) = 14 then digits else null
                        )
                    ),

            allCandidates =
                List.Distinct(
                    List.Select(
                        fmtCandidates & rawCandidates,
                        each _ <> null and _ <> KnownDigits and List.Count(List.Distinct(Text.ToList(_))) > 1
                    )
                ),

            valid = List.Select(allCandidates, each IsValidCNPJDigits(_)),
            picked =
                if List.Count(valid) > 0 then valid{0}
                else if List.Count(allCandidates) > 0 then allCandidates{0}
                else null
        in
            FormatCNPJ(picked),

    SliceBetween = (txt as text, startMarkers as list, endMarkers as list) as nullable text =>
        let
            startRec = FindFirstMarkerRecord(txt, startMarkers),
            startPos = if startRec = null then null else startRec[Pos] + Text.Length(startRec[Marker]),
            txtAfter = if startPos = null then null else Text.Middle(txt, startPos),

            endFound =
                if txtAfter = null then {}
                else
                    List.RemoveNulls(
                        List.Transform(
                            endMarkers,
                            each
                                let p = Text.PositionOf(txtAfter, _, Occurrence.First)
                                in if p >= 0 then p else null
                        )
                    ),

            endPos = if List.Count(endFound) = 0 then null else List.Min(endFound),

            result =
                if txtAfter = null then null
                else if endPos = null then txtAfter
                else Text.Start(txtAfter, endPos)
        in
            result,

    ExtractCNPJByLabeledWindow = (t as nullable text, labels as list, optional windowSize as nullable number) as nullable text =>
        let
            U = NormalizeKeepSpace(t),
            W = if windowSize = null then 250 else windowSize,
            labelRec = FindFirstMarkerRecord(U, labels),
            seg = if labelRec = null then null else Text.Middle(U, labelRec[Pos], W),
            result = if seg = null then null else FirstValidCNPJInText(seg)
        in
            result,

    ExtractCNPJPrestador = (t as nullable text) as nullable text =>
        let
            U = NormalizeKeepSpace(t),

            blocoPrestador1 =
                SliceBetween(
                    U,
                    {"PRESTADOR DE SERVIÇOS", "PRESTADOR DE SERVICOS"},
                    {"TOMADOR DE SERVIÇOS", "TOMADOR DE SERVICOS", "TOMADOR DO SERVIÇO", "TOMADOR DO SERVICO", "DISCRIMINAÇÃO DOS SERVICOS", "DISCRIMINACAO DOS SERVICOS", "DESCRIÇÃO DO SERVIÇO", "DESCRICAO DO SERVICO"}
                ),

            blocoPrestador2 =
                SliceBetween(
                    U,
                    {"EMITENTE DA NFS-E", "EMITENTE DA NFS-E", "EMITENTE DA NFS-e", "EMITENTE"},
                    {"TOMADOR DO SERVIÇO", "TOMADOR DO SERVICO", "TOMADOR DE SERVIÇOS", "TOMADOR DE SERVICOS", "DESCRIÇÃO DO SERVIÇO", "DESCRICAO DO SERVICO", "DETALHAMENTO DO SERVIÇO", "DETALHAMENTO DO SERVICO"}
                ),

            blocoFornecedor =
                SliceBetween(
                    U,
                    {"FORNECEDOR", "DADOS DO FORNECEDOR", "DADOSDOFORNECEDOR", "EMISSOR"},
                    {"TOMADOR", "CLIENTE", "DESTINATÁRIO", "DESTINATARIO", "PAGADOR", "SACADO"}
                ),

            porJanelaPrestador =
                ExtractCNPJByLabeledWindow(
                    U,
                    {
                        "PRESTADOR DE SERVIÇOS", "PRESTADOR DE SERVICOS",
                        "EMITENTE DA NFS-E", "EMITENTE DA NFS-e", "EMITENTE",
                        "DADOS DO PRESTADOR", "DADOS DO EMITENTE", "EMISSOR", "FORNECEDOR"
                    },
                    300
                ),

            cand1 = if blocoPrestador1 <> null then FirstValidCNPJInText(blocoPrestador1) else null,
            cand2 = if cand1 = null and blocoPrestador2 <> null then FirstValidCNPJInText(blocoPrestador2) else null,
            cand3 = if cand1 = null and cand2 = null and blocoFornecedor <> null then FirstValidCNPJInText(blocoFornecedor) else null,

            result =
                if cand1 <> null then cand1
                else if cand2 <> null then cand2
                else if cand3 <> null then cand3
                else porJanelaPrestador
        in
            result,

    // =========================
    // CNPJ NO NOME DO ARQUIVO (SE EXPLÍCITO)
    // NÃO CONCATENA PEDIDO + NF
    // =========================
    CNPJFromName = (name as nullable text) as nullable text =>
        let
            base =
                if name = null then null
                else if Text.EndsWith(name, ".pdf") or Text.EndsWith(name, ".PDF")
                then Text.Start(name, Text.Length(name)-4)
                else name,

            upperBase = if base = null then null else Text.Upper(base),

            fmtCandidates =
                if upperBase = null or Text.Length(upperBase) < 18 then {}
                else
                    List.RemoveNulls(
                        List.Transform(
                            {0..(Text.Length(upperBase)-18)},
                            (i) =>
                                let
                                    s18 = Text.Middle(upperBase, i, 18),
                                    digits = OnlyDigits(s18),
                                    isFmt =
                                        Text.Middle(s18, 2, 1) = "." and
                                        Text.Middle(s18, 6, 1) = "." and
                                        Text.Middle(s18, 10, 1) = "/" and
                                        Text.Middle(s18, 15, 1) = "-" and
                                        digits <> null and
                                        Text.Length(digits) = 14
                                in
                                    if isFmt then digits else null
                        )
                    ),

            parts =
                if base = null
                then {}
                else
                    List.Select(
                        Text.SplitAny(base, " _-().[]{}"),
                        each _ <> null and Text.Length(Text.Trim(_)) > 0
                    ),

            digit14Parts =
                List.Transform(parts, each OnlyDigits(_)),

            rawCandidates =
                List.Select(
                    digit14Parts,
                    each _ <> null and Text.Length(_) = 14
                ),

            allCandidates =
                List.Distinct(
                    List.Select(
                        fmtCandidates & rawCandidates,
                        each _ <> null
                            and _ <> OnlyDigits(KnownTomadorCNPJ)
                            and List.Count(List.Distinct(Text.ToList(_))) > 1
                            and IsValidCNPJDigits(_)
                    )
                ),

            picked =
                if List.Count(allCandidates) > 0 then allCandidates{0} else null
        in
            FormatCNPJ(picked),

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
            cleanedAfter =
                if afterMarker = null then null
                else
                    let
                        t1 = Text.TrimStart(afterMarker, {" ","_","-",".",":","(",")","0".."9"}),
                        t2 = if Text.Length(t1) >= 2 and Text.Upper(Text.Start(t1, 2)) = "E " then Text.Range(t1, 2) else t1
                    in
                        Text.Trim(t2),

            parts = Text.Split(base, "_"),
            rem1  = if List.Count(parts) >= 1 then List.Skip(parts, 1) else parts,
            rem2 =
                if List.Count(rem1) = 0 then rem1
                else
                    let
                        p0 = rem1{0}?,
                        p0d = if p0 = null then "" else Text.Select(p0, {"0".."9"})
                    in
                        if Text.Length(p0d) = 10 then List.Skip(rem1, 1) else rem1,
            rem3 =
                if List.Count(rem2) = 0 then rem2
                else
                    let
                        p1 = rem2{0}?,
                        p1d = if p1 = null then "" else Text.Select(p1, {"0".."9"})
                    in
                        if Text.Length(p1d) >= 2 and Text.Length(p1d) <= 9 then List.Skip(rem2, 1) else rem2,
            joined = Text.Trim(Text.Combine(rem3, " ")),
            fornecedor =
                if cleanedAfter <> null and cleanedAfter <> "" then cleanedAfter
                else if joined = "" then null
                else joined
        in
            fornecedor,

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
                                    let
                                        tail = Text.Middle(base, pos + Text.Length(m))
                                    in
                                        GetDigitsAfter(tail)
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

    ExtractPedidoFromText = (t as nullable text) as nullable text =>
        let
            U = if t = null then "" else NormalizeKeepSpace(t),
            markers = {
                "PEDIDO DE COMPRA",
                "PEDIDO",
                "ORDEM DE COMPRA",
                "NUMERO DO PEDIDO",
                "NÚMERO DO PEDIDO"
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

    // NF SOMENTE PELO NOME DO ARQUIVO
    AddNF = Table.AddColumn(RemPedidoAux, "NUM_NF", each NFFromName([Name]), type text),

    // CNPJ CORRIGIDO
    AddCNPJ = Table.AddColumn(
        AddNF,
        "CNPJ",
        each
            let
                cnpjPrestador = ExtractCNPJPrestador([TextAll]),
                cnpjTexto = ExtractCNPJ([TextAll]),
                cnpjTextoLoose = ExtractCNPJLoose([TextAll]),
                cnpjNome = CNPJFromName([Name])
            in
                if cnpjPrestador <> null then cnpjPrestador
                else if cnpjTexto <> null then cnpjTexto
                else if cnpjTextoLoose <> null then cnpjTextoLoose
                else cnpjNome,
        type text
    ),

    AddForn = Table.AddColumn(AddCNPJ, "FORNECEDOR", each FornecedorFromName([Name]), type text),

    AddEmiss_Txt = Table.AddColumn(AddForn, "EMISSAO_TXT", each ExtractEmissao([TextAll]), type date),
    AddEmiss = Table.AddColumn(AddEmiss_Txt, "EMISSAO", each if [EMISSAO_TXT] <> null then [EMISSAO_TXT] else EmissaoFromName([Name]), type date),
    RemEmiss_Txt = Table.RemoveColumns(AddEmiss, {"EMISSAO_TXT"}),

    AddTipo = Table.AddColumn(RemEmiss_Txt, "TIPO", each DetectTipo([Name], [TextPg1], [TextAll]), type text),

    // >>> VALOR corrigido: tenta página 1 e, se falhar, tenta texto completo
    AddValor = Table.AddColumn(
        AddTipo,
        "VALOR",
        each
            let
                v1 = ExtractValor([TextPg1]),
                v2 = if v1 = null then ExtractValor([TextAll]) else v1
            in
                v2,
        type number
    ),

    AddNServ = Table.AddColumn(AddValor, "NUM_SERVICO", each ExtractNumServico([TextPg1]), type text),

    Selected =
        Table.SelectColumns(
            AddNServ,
            {"Name","PEDIDO","TIPO","NUM_NF","FORNECEDOR","CNPJ","EMISSAO","VALOR","NUM_SERVICO","TextPg1_NoSpace"}
        ),

    FinalRenamed = Table.RenameColumns(Selected, {{"TextPg1_NoSpace","TEXTO_PG1_SEM_ESPACO"}}),

    Final =
        if KeepOnlyWithPedido
        then Table.SelectRows(FinalRenamed, each [PEDIDO] <> null)
        else FinalRenamed,

    #"Linhas Filtradas" = Table.SelectRows(Final, each [PEDIDO] <> null),
    #"Colunas Removidas" = Table.RemoveColumns(#"Linhas Filtradas", {"NUM_SERVICO"}),
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Colunas Removidas",{{"EMISSAO", type date}})
in
    #"Tipo Alterado"
