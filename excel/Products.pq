let
    FolderPath = Text.From(Excel.CurrentWorkbook(){[Name="DataFolder"]}[Content]{0}[Column1]),
    Files = Table.SelectRows(Folder.Contents(FolderPath), each List.Contains({"products_01.csv.gz","products_02.csv.gz","products_03.csv.gz","products_04.csv.gz"}, [Name])),
    CheckFiles = if Table.RowCount(Files) <> 4 then error "В папке data должны находиться все четыре файла products_01.csv.gz ... products_04.csv.gz" else Files,
    ReadFiles = Table.AddColumn(CheckFiles, "Rows", each Table.PromoteHeaders(Csv.Document(Binary.Decompress([Content], Compression.GZip), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]), [PromoteAllScalars=true])),
    Combined = Table.Combine(ReadFiles[Rows]),
    KeepColumns = Table.SelectColumns(Combined, {"source_row","cat1","cat2","cat3","cat4","name","url","seller","brand","scheme","revenue","price","availability","drr"}),
    Typed = Table.TransformColumnTypes(KeepColumns, {{"source_row",Int64.Type},{"revenue",type number},{"price",type number},{"availability",type number},{"drr",type number}}, "en-US"),
    RequiredValues = if Table.RowCount(Table.SelectRows(Typed, each [revenue]=null or [drr]=null or [cat4]=null or [revenue]<0 or [drr]<0)) > 0 then error "Проверьте обязательные поля и отрицательные значения" else Typed,
    Million = Table.AddColumn(RequiredValues, "million", each if [revenue] > 1000000 then 1 else 0, Int64.Type),
    AdSpend = Table.AddColumn(Million, "ad_spend_est", each [revenue]*[drr], type number)
in
    AdSpend
