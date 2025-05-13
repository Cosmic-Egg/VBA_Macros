let
    // 1. Get files
    Source = Folder.Files("C:\Your\Folder\Path"),  // <-- update path
    FilterXLSM = Table.SelectRows(Source, each Text.EndsWith([Extension], ".xlsm")),

    // 2. Extract "Benefit Input" sheet safely
    AddSheet = Table.AddColumn(FilterXLSM, "Sheet", each try 
        Table.SelectRows(Excel.Workbook([Content], true), each [Item] = "Benefit Input" and [Kind] = "Sheet"){0}[Data] 
        otherwise null),

    // 3. Filter out nulls
    Filtered = Table.SelectRows(AddSheet, each [Sheet] <> null),

    // 4. Clean and prep each file
    TransformEach = Table.AddColumn(Filtered, "Cleaned", each 
        let
            FileName = Text.BeforeDelimiter([Name], ".xlsm"),
            Original = [Sheet],
            RemovedTop = Table.Skip(Original, 7),
            RemovedFirstCols = Table.RemoveColumns(RemovedTop, List.FirstN(Table.ColumnNames(RemovedTop), 7)),
            // Determine columns that contain any non-null value
            NonBlankCols = List.Select(
                Table.ColumnNames(RemovedFirstCols),
                (col) => List.NonNullCount(RemovedFirstCols[col]) > 0
            ),
            CleanedCols = Table.SelectColumns(RemovedFirstCols, NonBlankCols),
            WithIndex = Table.AddIndexColumn(CleanedCols, "RowIndex", 1, 1, Int64.Type),
            Renamed = Table.RenameColumns(WithIndex, List.Transform(
                List.RemoveItems(Table.ColumnNames(WithIndex), {"RowIndex"}), 
                each {_, FileName & "_" & _}
            ))
        in
            Renamed
    ),

    // 5. Get cleaned tables
    TablesToMerge = TransformEach[Cleaned],

    // 6. Merge all tables by RowIndex
    Merged = List.Accumulate(
        List.Skip(TablesToMerge, 1),
        List.First(TablesToMerge),
        (state, current) => Table.Join(state, "RowIndex", current, "RowIndex", JoinKind.Inner)
    ),

    // 7. Remove helper index
    Final = Table.RemoveColumns(Merged, {"RowIndex"})
in
    Final
