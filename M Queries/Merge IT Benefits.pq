let
    // 1. Load all .xlsm files from folder
    Source = Folder.Files("C:\Your\Folder\Path"),  //  UPDATE path
    XLSM_Files = Table.SelectRows(Source, each Text.EndsWith([Extension], ".xlsm")),

    // 2. Extract the "Benefit Input" sheet from each file
    AddSheet = Table.AddColumn(XLSM_Files, "SheetData", each try 
        Table.SelectRows(Excel.Workbook([Content], true), each [Item] = "Benefit Input" and [Kind] = "Sheet"){0}[Data] 
        otherwise null),

    // 3. Remove files where sheet wasn't found
    ValidSheets = Table.SelectRows(AddSheet, each [SheetData] <> null),

    // 4. Clean each sheet: skip rows, skip columns, filter non-blank columns
    CleanedSheets = Table.AddColumn(ValidSheets, "Cleaned", each 
        let
            FileName = Text.BeforeDelimiter([Name], ".xlsm"),
            Sheet = [SheetData],
            SkipRows = Table.Skip(Sheet, 7),
            SkipCols = Table.RemoveColumns(SkipRows, List.FirstN(Table.ColumnNames(SkipRows), 7)),
            // Identify columns with any non-null value
            ColumnsToKeep = List.Select(
                Table.ColumnNames(SkipCols),
                (cName) => List.NonNullCount(SkipCols[cName]) > 0
            ),
            Trimmed = Table.SelectColumns(SkipCols, ColumnsToKeep),
            WithIndex = Table.AddIndexColumn(Trimmed, "RowIndex", 1, 1, Int64.Type),
            Renamed = Table.RenameColumns(WithIndex, List.Transform(
                List.RemoveItems(Table.ColumnNames(WithIndex), {"RowIndex"}),
                each {_, FileName & "_" & _}
            ))
        in
            Renamed
    ),

    // 5. Get only the cleaned tables
    TableList = CleanedSheets[Cleaned],

    // 6. Merge side-by-side on RowIndex
    Merged = List.Accumulate(
        List.Skip(TableList, 1),
        List.First(TableList),
        (state, current) => Table.Join(state, "RowIndex", current, "RowIndex", JoinKind.Inner)
    ),

    // 7. Remove RowIndex
    Final = Table.RemoveColumns(Merged, {"RowIndex"})
in
    Final
