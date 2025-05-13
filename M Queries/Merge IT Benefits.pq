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

    // ✅ Convert first two rows to lists
    Row1 = Record.ToList(Cleaned{0}),
    Row2 = Record.ToList(Cleaned{1}),

    // ✅ Update Row1: if Row2 = "OON", use previous column’s Row1
    ModifiedRow1 = List.Transform(
        List.Zip({Row1, Row2, {null} & Row1}),
        each if _{1} = "OON" then _{2} else _{0}
    ),

    // ✅ Build new top 2 rows
    NewTop2 = Table.FromRows({ModifiedRow1, Row2}, Table.ColumnNames(Cleaned)),

    // ✅ Combine with the remaining rows
    RestRows = Table.Skip(Cleaned, 2),
    Recombined = Table.Combine({NewTop2, RestRows}),

    // ✅ Get new list of column names
    ColNames = Table.ColumnNames(Recombined),

    // ✅ Identify columns that match pattern HXXXX-XXX-XXX
    ValidCols = List.Select(
        List.Zip({ColNames, ModifiedRow1}),
        each 
            let 
                val = _{1},
                isText = (val <> null and Text.Length(val) = 13),
                match = isText and Text.Start(val, 1) = "H" and
                        Text.Middle(val, 5, 1) = "-" and
                        Text.Middle(val, 9, 1) = "-"
            in
                match
    ),
    FilteredCols = List.Transform(ValidCols, each _{0}),
    FinalFiltered = Table.SelectColumns(Recombined, FilteredCols),

    // ✅ Add row index and rename columns
    WithIndex = Table.AddIndexColumn(FinalFiltered, "RowIndex", 1, 1, Int64.Type),
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
