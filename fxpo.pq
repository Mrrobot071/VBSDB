let

    // =====================================================

    // CAMINHO DA PASTA

    // =====================================================

    CaminhoPastaPDFs = "C:\Users\lucas.nathan\OneDrive - ALLOS\Área de Trabalho\Anexostudo",



    // =====================================================

    // HELPERS

    // =====================================================

    nzText = (t as nullable text) as text =>

        if t = null then "" else Text.From(t),



    fnNormalizeSoft = (t as nullable text) as text =>

        if t = null then ""

        else

            let

                s0 = Text.Clean(t),

                s1 = Text.Replace(Text.Replace(Text.Replace(s0, "#(lf)", " "), "#(cr)", " "), "#(tab)", " "),

                s2 = Text.Replace(s1, Character.FromNumber(160), " "),

                tokens = List.Select(Text.Split(s2, " "), each _ <> "")

            in

                Text.Combine(tokens, " "),



    fnDataBlockToTextRaw = (dataBlock as any) as text =>

        if (dataBlock = null) or (not (dataBlock is table)) then nzText(dataBlock)

        else

            let

                rows = Table.ToRows(dataBlock),

                rowTexts = List.Transform(rows, each Text.Combine(List.Transform(_, nzText), " ")),

                out = Text.Combine(rowTexts, "#(lf)")

            in

                out,



    fnGetTextoPagina1Raw = (pdfTables as nullable table) as text =>

        if pdfTables = null then ""

        else

            let

                pag1 =

                    if Table.HasColumns(pdfTables, "Page")

                    then Table.SelectRows(pdfTables, each [Page] = 1)

                    else Table.FirstN(pdfTables, 1),

                blocos = try Table.Column(pag1, "Data") otherwise {},

                out = Text.Combine(List.Transform(blocos, fnDataBlockToTextRaw), "#(lf)")

            in

                out,



    fnTextNoSpacesKeepDigits = (t as nullable text) as text =>

        if t = null then ""

        else

            Text.Remove(

                t,

                {" ", "#(lf)", "#(cr)", "#(tab)", "|", "¦", "｜", Character.FromNumber(160)}

            ),



    // =====================================================

    // ✅ DETECÇÃO DEFINITIVA DE ORDEM DE COMPRA

    // =====================================================

    fnTemOrdemCompra = (t as nullable text) as logical =>

        let

            txt = Text.Lower(fnNormalizeSoft(t)),

            compacto = Text.Remove(txt, {" ", "#(lf)", "#(cr)", "#(tab)"}),

            resultado =

                Text.Contains(compacto, "ordemdecompra")

                or Text.Contains(compacto, "ordemdecomprapo")

                or Text.Contains(compacto, "(po)")

        in

            resultado,



    // =====================================================

    // PEDIDO – começa com 45

    // =====================================================

    fnExtractPedidoApenasTexto = (textoPdf as nullable text) as nullable text =>

        let

            textoLower = Text.Lower(nzText(textoPdf)),

            dig =

                Text.Combine(

                    List.Transform(

                        Text.ToList(fnNormalizeSoft(textoLower)),

                        each if List.Contains({"0".."9"}, _) then _ else " "

                    ),

                    ""

                ),

            tokens = List.Select(Text.Split(dig, " "), each Text.StartsWith(_, "45")),

            resultado = if List.Count(tokens) > 0 then Text.Start(tokens{0}, 10) else null

        in

            resultado,



    // =====================================================

    // ✅ SAP – 10 DÍGITOS APÓS "CÓDIGO:" (PRIORIZA "0010")

    //     * Preserva zeros à esquerda na extração

    //     * Depois remove SOMENTE o "00" imediatamente antes de "10"

    //     * Retorna como TEXTO (evita perder zeros)

    // =====================================================

    fnExtractSAP10dAposCodigo = (t as nullable text) as nullable text =>

        if t = null then null

        else

            let

                texto = fnNormalizeSoft(t),

                low = Text.Lower(texto),



                // Posições possíveis do rótulo

                posCod2 = Text.PositionOf(low, "código:"),

                posCod1 = Text.PositionOf(low, "codigo:"),

                posCod4 = Text.PositionOf(low, "código"),

                posCod3 = Text.PositionOf(low, "codigo"),

                listaPos = List.RemoveItems({posCod2, posCod1, posCod4, posCod3}, {-1}),

                pos = if List.Count(listaPos) > 0 then List.Min(listaPos) else -1,



                // Trecho após o rótulo

                trechoRaw = if pos < 0 then null else Text.Middle(texto, pos, 200),

                posDoisPontos = if trechoRaw <> null then Text.PositionOf(trechoRaw, ":") else -1,

                afterLabel =

                    if trechoRaw = null then null

                    else if posDoisPontos >= 0 then Text.Middle(trechoRaw, posDoisPontos + 1)

                    else trechoRaw,



                // Mantém somente dígitos

                onlyDigits =

                    if afterLabel = null then ""

                    else

                        Text.Combine(

                            List.Transform(

                                Text.ToList(afterLabel),

                                each if List.Contains({"0".."9"}, _) then _ else ""

                            ),

                            ""

                        ),



                // Todas as janelas de 10 dígitos consecutivos

                candAll =

                    if Text.Length(onlyDigits) >= 10 then

                        List.Generate(

                            () => 0,

                            each _ <= Text.Length(onlyDigits) - 10,

                            each _ + 1,

                            each Text.Range(onlyDigits, _, 10)

                        )

                    else

                        {},



                // Prioriza as que começam com "0010"

                cand0010 = List.Select(candAll, each Text.StartsWith(_, "0010")),

                sap10 = if List.Count(cand0010) > 0 then cand0010{0}

                        else if List.Count(candAll) > 0 then candAll{0}

                        else null,



                // 👇 Remove APENAS o "00" imediatamente antes de "10"

                sapFinal =

                    if sap10 <> null and Text.StartsWith(sap10, "0010")

                    then "10" & Text.Range(sap10, 4)   // vira 8 dígitos: "10xxxxxx"

                    else sap10                          // mantém como veio (10 dígitos)

            in

                sapFinal,



    // =====================================================

    // CNPJ

    // =====================================================

    fnExtractCNPJComMascara = (t as nullable text) as nullable text =>

        if t = null then null

        else

            let

                txt = Text.Upper(fnNormalizeSoft(t)),

                cand =

                    List.RemoveNulls(

                        List.Transform(

                            List.Generate(

                                () => 0,

                                each _ <= Text.Length(txt) - 18,

                                each _ + 1,

                                each Text.Middle(txt, _, 18)

                            ),

                            each

                                if Text.Length(_) = 18

                                   and Text.Middle(_, 2, 1) = "."

                                   and Text.Middle(_, 6, 1) = "."

                                   and Text.Middle(_, 10, 1) = "/"

                                   and Text.Middle(_, 15, 1) = "-"

                                then _

                                else null

                        )

                    )

            in

                if List.Count(cand) > 0 then cand{0} else null,



    // =====================================================

    // VALOR

    // =====================================================

    fnExtractValorFromPag1 = (t as nullable text) as nullable number =>

        if t = null then null

        else

            let

                txt = fnNormalizeSoft(t),

                chars = Text.ToList(txt),

                limpo =

                    Text.Combine(

                        List.Transform(

                            chars,

                            each if List.Contains({"0".."9", ".", ",", " "}, _) then _ else " "

                        ),

                        ""

                    ),

                tokens = List.Select(Text.Split(limpo, " "), each Text.Contains(_, ",")),

                vals = List.RemoveNulls(List.Transform(tokens, each try Number.FromText(_, "pt-BR") otherwise null))

            in

                if List.Count(vals) > 0 then List.Max(vals) else null,



    // =====================================================

    // PIPELINE

    // =====================================================

    Fonte = Folder.Files(CaminhoPastaPDFs),

    ApenasPDF = Table.SelectRows(Fonte, each Text.Lower([Extension]) = ".pdf"),

    PdfLido = Table.AddColumn(ApenasPDF, "Pdf", each try Pdf.Tables([Content]) otherwise null),

    TextoRaw = Table.AddColumn(PdfLido, "TEXTO_PAG1_RAW", each fnGetTextoPagina1Raw([Pdf])),



    // ✅ Filtro definitivo – detecta O.C.

    FiltrarOrdemCompra = Table.SelectRows(TextoRaw, each fnTemOrdemCompra([TEXTO_PAG1_RAW])),



    AddPedido = Table.AddColumn(FiltrarOrdemCompra, "PEDIDO", each fnExtractPedidoApenasTexto([TEXTO_PAG1_RAW])),



    // ✅ COD SAP já com remoção do "00" antes do "10" quando aplicável

    AddSAP = Table.AddColumn(AddPedido, "COD SAP", each fnExtractSAP10dAposCodigo([TEXTO_PAG1_RAW]), type text),



    AddCNPJ = Table.AddColumn(AddSAP, "CNPJ", each fnExtractCNPJComMascara([TEXTO_PAG1_RAW])),

    AddValor = Table.AddColumn(AddCNPJ, "VALOR", each fnExtractValorFromPag1([TEXTO_PAG1_RAW])),

    AddTextoResumo = Table.AddColumn(AddValor, "TEXTO_LIMPO", each fnTextNoSpacesKeepDigits([TEXTO_PAG1_RAW])),



    ResultadoFinal =

        Table.SelectColumns(

            AddTextoResumo,

            {"PEDIDO", "COD SAP", "CNPJ", "VALOR", "TEXTO_LIMPO", "Name"}

        ),



    // 👉 Mantemos "COD SAP" como texto (para evitar perda de zeros em outros casos)

    #"Tipo Alterado" =

        Table.TransformColumnTypes(

            ResultadoFinal,

            {

                {"PEDIDO", Int64.Type},

                {"VALOR", Currency.Type},

                {"CNPJ", type text},

                {"COD SAP", type text}

            }

        )



in

    #"Tipo Alterado"
