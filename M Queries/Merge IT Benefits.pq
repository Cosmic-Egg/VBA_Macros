let
    // STEP 1: Load all .xlsm files from the folder
    Source = Folder.Files("C:\Your\Folder\Path"),  // 🔁 Change to your path
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

            // Skip first 7 rows
            SkippedRows = Table.Skip(RawTable, 7),

            // Remove first 7 columns
            RemainingCols = Table.RemoveColumns(SkippedRows, List.FirstN(Table.ColumnNames(SkippedRows), 7)),

            // Clean out columns that are entirely blank
            NonBlankCols = List.Select(
                Table.ColumnNames(RemainingCols),
                (colName) =>
                    let
                        colValues = try Table.Column(RemainingCols, colName) otherwise {},
                        nonNullValues = List.Select(colValues, each _ <> null and _ <> "")
                    in
                        List.Count(nonNullValues) > 0
            ),
            Filtered = Table.SelectColumns(RemainingCols, NonBlankCols),

            // Add an index for horizontal joining
            WithIndex = Table.AddIndexColumn(Filtered, "RowIndex", 1, 1, Int64.Type),

            // Rename columns to include filename
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

    // STEP 6: Remove the RowIndex column
    Final = Table.RemoveColumns(Merged, {"RowIndex"})
in
    Final
