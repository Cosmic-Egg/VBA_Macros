

   
let
    ParamTable = Table.FromRecords({
        [tType = "All", mType = "GMD"]
    }),
    
    pTable = ParamTable,
    firstRow = pTable{0},
    //filteredValue = Record.Field(firstRow,"tType"),
    Source = Csv.Document(File.Contents("filepath.csv"),[Delimiter=",", Columns=4, Encoding=1252, QuoteStyle=QuoteStyle.None]),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true])
    /*filteredData =
    if filteredValue = "All" or filteredValue = null then
        PromotedHeaders
    else
        Table.SelectRows(PromotedHeaders, each [transactionType] = filteredValue)
    */
in
    //filteredData
    firstRow