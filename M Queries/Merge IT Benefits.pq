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
        ModifiedRow1 = List.Transform(
            List.Zip({Row1, Row2, {null} & Row1}),
            each if _{1} = "OON" then _{2} else _{0}
        ),

        // ✅ Build new top two rows
        NewTop2 = Table.FromRows({ModifiedRow1, Row2}, ColNames),

        // ✅ Combine with the rest of the data
        RestRows = Table.Skip(Cleaned, 2),
        Recombined = Table.Combine({NewTop2, RestRows}),

        // ✅ OPTIONAL: Add RowIndex and rename columns by file
        WithIndex = Table.AddIndexColumn(Recombined, "RowIndex", 1, 1, Int64.Type),
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
