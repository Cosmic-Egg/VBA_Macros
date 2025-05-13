let
    // 1. Load all files from the folder
    Source = Folder.Files("C:\Your\Folder\Path"),  // <--  Change this path
    FilterXLSM = Table.SelectRows(Source, each Text.EndsWith([Extension], ".xlsm")),

    // 2. Extract the "Benefit Input" sheet from each file
    AddSheet = Table.AddColumn(FilterXLSM, "Sheet", each try 
        Table.SelectRows(Excel.Workbook([Content], true), each [Item] = "Benefit Input" and [Kind] = "Sheet"){0}[Data] 
        otherwise null),

    // 3. Remove any rows without the target sheet
    Filtered = Table.SelectRows(AddSheet, each [Sheet] <> null),

    // 4. Remove first 7 rows, then first 7 columns, and remove blank columns based on last row
    TransformEach = Table.AddColumn(Filtered, "Cleaned", each 
        let
            Original = [Sheet],
            RemovedTop = Table.Skip(Original, 7),
            RemovedCols = Table.RemoveColumns(RemovedTop, List.FirstN(Table.ColumnNames(RemovedTop), 7)),
            LastRow = Table.LastN(RemovedCols, 1){0},
            NonBlankCols = List.Select(Table.ColumnNames(RemovedCols), each Record.Field(LastRow, _) <> null and Text.Trim(Text.From(Record.Field(LastRow, _))) <> ""),
            Cleaned = Table.SelectColumns(RemovedCols, NonBlankCols),
            WithIndex = Table.AddIndexColumn(Cleaned, "RowIndex", 1, 1, Int64.Type),
            Renamed = Table.RenameColumns(WithIndex, List.Transform(List.RemoveItems(Table.ColumnNames(WithIndex), {"RowIndex"}), each {_, Text.BeforeDelimiter([Name], ".xlsm") & "_" & _}))
        in
            Renamed
    ),

    // 5. Get the list of cleaned tables
    TablesToMerge = TransformEach[Cleaned],

    // 6. Merge all tables by RowIndex
    Merged = List.Accumulate(
        List.Skip(TablesToMerge, 1),
        List.First(TablesToMerge),
        (state, current) => Table.Join(state, "RowIndex", current, "RowIndex", JoinKind.Inner)
    ),

    // 7. Remove RowIndex column
    Final = Table.RemoveColumns(Merged, {"RowIndex"})
in
    Final
