let
    // STEP 1: Load all .xlsm files from the folder
    Source = Folder.Files("C:\Your\Folder\Path"),  // 🔁 Change this path
    XLSMFiles = Table.SelectRows(Source, each Text.EndsWith([Extension], ".xlsm")),

    // STEP 2: Extract the 'Benefit Input' sheet from each file
    ExtractedSheets = Table.AddColumn(XLSMFiles, "Sheet", each try 
        Table.SelectRows(Excel.Workbook([Content], true), each [Item] = "Benefit Input" and [Kind] = "Sheet"){0}[Data]
        otherwise null),

    // STEP 3: Remove files where the sheet wasn't found
    ValidSheets = Table.SelectRows(ExtractedSheets, each [Sheet] <> null),

    // STEP 4: Clean each sheet
    ProcessedSheets = Table.AddColumn(ValidSheets, "Processed", each 
    let
        FileName = Text.BeforeDelimiter([Name], ".xlsm"),
        RawTable = [Sheet],

        // ✅ Skip first 7 header rows, then take first 242 data rows
        SkippedTop = Table.Skip(RawTable, 7),
        TrimmedRows = Table.FirstN(SkippedTop, 242),

        // ✅ Remove first 7 columns
        RemainingCols = Table.RemoveColumns(TrimmedRows, List.FirstN(Table.ColumnNames(TrimmedRows), 7)),

        // ✅ Replace Excel errors (like #N/A) with nulls
        Cleaned = Table.ReplaceErrorValues(RemainingCols, List.Transform(Table.ColumnNames(RemainingCols), each {_, null})),

        // ✅ Get number of columns to ensure safe padding
        ColNames = Table.ColumnNames(Cleaned),
        ColCount = List.Count(ColNames),

        // ✅ Get top two rows as lists
        Top2Rows = Table.ToRows(Table.FirstN(Cleaned, 2)),
        RawRow1 = if List.Count(Top2Rows) > 0 then Top2Rows{0} else {},
        RawRow2 = if List.Count(Top2Rows) > 1 then Top2Rows{1} else {},

        // ✅ Pad both rows with nulls if needed
        PadRow = (row as list, count as number) =>
            if List.Count(row) < count then
                List.Combine({row, List.Repeat({null}, count - List.Count(row))})
            else
                row,

        Row1 = PadRow(RawRow1, ColCount),
        Row2 = PadRow(RawRow2, ColCount),

        // ✅ Apply transformation: if Row2 = "OON", fill Row1 from previous Row1 column
        ModifiedRow1Full = List.Transform(
            List.Zip({Row1, Row2, {null} & Row1}),
            each 
                let
                    val1 = _{0},
                    val2 = _{1},
                    prev = _{2}
                in
                    if Text.From(val2) = "OON" then prev else val1
        ),

        // ✅ Truncate to match exact column count
        ModifiedRow1 = List.FirstN(ModifiedRow1Full, ColCount),

        // ✅ Remove columns where the value in Row 1 doesn't match H####-###-###
        ValidColumnNames = List.Select(
            List.Zip({ModifiedRow1, ColNames}),
            each Text.Length(Text.From(_{0})) = 12 and Text.Start(Text.From(_{0}), 1) = "H" and Text.Middle(Text.From(_{0}), 1, 4) = "H" and Text.Middle(Text.From(_{0}), 5, 4) = "-" and Text.Middle(Text.From(_{0}), 10, 3) = "-" and Text.Middle(Text.From(_{0}), 7, 3) <> null
        ),

        ValidColumnNamesOnly = List.Transform(ValidColumnNames, each _{1}),

        // ✅ Filter columns that don't match the pattern in Row 1 (H####-###-###)
        FilteredColumns = Table.SelectColumns(Cleaned, ValidColumnNamesOnly),

        // ✅ OPTIONAL: Add RowIndex and rename columns by file
        WithIndex = Table.AddIndexColumn(FilteredColumns, "RowIndex", 1, 1, Int64.Type),
        ColumnsToRename = List.RemoveItems(Table.ColumnNames(WithIndex), {"RowIndex"}),
        RenameMap = List.Transform(ColumnsToRename, each {_, FileName & "_" & _}),
        Renamed = Table.RenameColumns(WithIndex, RenameMap, MissingField.Ignore)
    in
        Renamed


    ),

    // STEP 5: Combine all processed tables horizontally
    TablesToJoin = ProcessedSheets[Processed],
    Merged = List.Accumulate(
        List.Skip(TablesToJoin, 1),
        List.First(TablesToJoin),
        (state, next) => Table.Join(state, "RowIndex", next, "RowIndex", JoinKind.Inner)
    ),

    // STEP 6: Remove index column
    Final = Table.RemoveColumns(Merged, {"RowIndex"})
in
    Final
