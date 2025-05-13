let
    // 1. Load all .xlsm files from folder
    Source = Folder.Files("C:\Your\Folder\Path"),  // 🔁 Update this
    XLSM_Files = Table.SelectRows(Source, each Text.EndsWith([Extension], ".xlsm")),

    // 2. Extract "Benefit Input" sheet from each file
    AddSheet = Table.AddColumn(XLSM_Files, "SheetData", each try 
        Table.SelectRows(Excel.Workbook([Content], true), each [Item] = "Benefit Input" and [Kind] = "Sheet"){0}[Data] 
        otherwise null),

    // 3. Keep only valid sheets
    ValidSheets = Table.SelectRows(AddSheet, each [SheetData] <> null),

    // 4. Clean each sheet
    CleanedSheets = Table.AddColumn(ValidSheets, "Cleaned", each 
        let
            FileName = Text.BeforeDelimiter([Name], ".xlsm"),
            Sheet = [SheetData],
            SkipRows = Table.Skip(Sheet, 7),
            SkipCols = Table.RemoveColumns(SkipRows, List.FirstN(Table.ColumnNames(SkipRows), 7)),

            // ✅ Filter columns that have at least one non-null value
            ColumnsToKeep = List.Select(
                Table.ColumnNames(SkipCols),
                (cName) => List.NonNullCount(List.Transform(Table.ToRows(SkipCols), each Record.Field(_, cName))) > 0
            ),

            Trimmed = Table.SelectColumns(SkipCols, ColumnsToKeep),

            // Add index to merge later
            WithIndex = Table.AddIndexColumn(Trimmed, "RowIndex", 1, 1, Int64.Type),

            // ✅ Rename columns safely
            ColumnsToRename = List.RemoveItems(Table.ColumnNames(WithIndex), {"RowIndex"}),
            RenamePairs = List.Transform(ColumnsToRename, each {_, FileName & "_" & _}),
            Renamed = if List.Count(RenamePairs) > 0 then Table.RenameColumns(WithIndex, RenamePairs) else WithIndex
        in
            Renamed
    ),

    // 5. Combine all cleaned tables
    TableList = CleanedSheets[Cleaned],

    // 6. Merge horizontally by RowIndex
    Merged = List.Accumulate(
        List.Skip(TableList, 1),
        List.First(TableList),
        (state, current) => Table.Join(state, "RowIndex", current, "RowIndex", JoinKind.Inner)
    ),

    // 7. Remove RowIndex
    Final = Table.RemoveColumns(Merged, {"RowIndex"})
in
    Final
