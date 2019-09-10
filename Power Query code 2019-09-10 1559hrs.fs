// GetValue
(rangeName,index) => Excel.CurrentWorkbook(){[Name=rangeName]}[Content]{index}[User Input]

// workbookSheet1(binary)
(Workbook as binary) =>
let
    #"Workbook" = Excel.Workbook(Workbook, null, true),
    Export_Sheet = #"Workbook"{[Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Export_Sheet, [PromoteAllScalars=true])
in
    #"Promoted Headers"

// StripPunctuation(text)
let
    StripPunctuation = (xString as text) =>
        let
            PunctuationString = "!""#(#)$%&''()*+, -./:;<=>?@[\]^_`{|}~",
            PunctuationList = Text.ToList(PunctuationString),
            Output = Text.Remove(xString,PunctuationList),
            Output2 = Text.Start(Output,40),
            Output3 = Text.Upper(Output2)
        in
            Output3 as text
in StripPunctuation

// CountryFilterExpand
(Source as table, Country as text) =>
let
    #"Filtered Rows" = Table.SelectRows(Source, each Text.StartsWith([Source.Name], Country)),
    #"Expanded workbookSheet1(binary)" = Table.ExpandTableColumn(#"Filtered Rows", "workbookSheet1(binary)", Table.ColumnNames(#"Filtered Rows"{0}[#"workbookSheet1(binary)"]))
in
    #"Expanded workbookSheet1(binary)"

// MasterSheet
let
    Source = Excel.Workbook(File.Contents(GetValue("Config",1)), null, true),
    Sheet1_Sheet = Source{[Item="Sheet1",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Sheet1_Sheet, [PromoteAllScalars=true]),
    #"Removed Other Columns" = Table.SelectColumns(#"Promoted Headers",{"Importer Company", "Standard Manufacturing CO", "Type", "Remarks"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Other Columns",{{"Importer Company", type text}, {"Standard Manufacturing CO", type text}, {"Type", type text}, {"Remarks", type text}}),
    #"Invoked Custom Function" = Table.AddColumn(#"Changed Type", "Importer_", each #"StripPunctuation(text)"([Importer Company])),
    #"Removed Duplicates" = Table.Distinct(#"Invoked Custom Function", {"Importer_", "Standard Manufacturing CO"})
in
    #"Removed Duplicates"

// MasterSheet duplicates
let
    Source = MasterSheet,
    #"Grouped Rows" = Table.Group(Source, {"Importer_"}, {{"Count", each Table.RowCount(_), type number}, {"Details", each _, type table}}),
    #"Filtered Rows" = Table.SelectRows(#"Grouped Rows", each [Count] > 1),
    #"Removed Other Columns" = Table.SelectColumns(#"Filtered Rows",{"Details"}),
    #"Expanded Details" = Table.ExpandTableColumn(#"Removed Other Columns", "Details", {"Importer Company", "Standard Manufacturing CO", "Type", "Remarks", "Importer_"}, {"Importer Company", "Standard Manufacturing CO", "Type", "Remarks", "Importer_.1"})
in
    #"Expanded Details"

// Raw data files
let
    Source = Folder.Files(GetValue("Config",0)),
    #"Filtered Hidden Files1" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true),
    #"Filtered Rows" = Table.SelectRows(#"Filtered Hidden Files1", each Text.StartsWith([Extension], ".xls")),
    #"Invoked Custom Function" = Table.AddColumn(#"Filtered Rows", "workbookSheet1(binary)", each #"workbookSheet1(binary)"([Content])),
    #"Removed Other Columns" = Table.SelectColumns(#"Invoked Custom Function",{"Name", "workbookSheet1(binary)"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Other Columns",{{"Name", "Source.Name"}})
in
    #"Renamed Columns"

// America
let
    Source = #"Raw data files",
    #"CountryFilterExpand" = CountryFilterExpand(Source,"America"),
    #"Invoked Custom Function" = Table.AddColumn(CountryFilterExpand, "Importer_", each #"StripPunctuation(text)"([FOREIGN IMPORTER NAME]))
in
    #"Invoked Custom Function"

// Britain
let
    Source = #"Raw data files",
    #"CountryFilterExpand" = CountryFilterExpand(Source,"Britain"),
    #"Invoked Custom Function" = Table.AddColumn(CountryFilterExpand, "Importer_", each #"StripPunctuation(text)"([Importer]))
in
    #"Invoked Custom Function"

// Canada
let
    Source = #"Raw data files",
    #"CountryFilterExpand" = CountryFilterExpand(Source,"Canada"),
    #"Invoked Custom Function" = Table.AddColumn(CountryFilterExpand, "Importer_", each #"StripPunctuation(text)"([IMPORTER NAME]))
in
    #"Invoked Custom Function"

// Denmark
let
    Source = #"Raw data files",
    #"CountryFilterExpand" = CountryFilterExpand(Source,"Denmark"),
    #"Invoked Custom Function" = Table.AddColumn(CountryFilterExpand, "Importer_", each #"StripPunctuation(text)"([Importer]))
in
    #"Invoked Custom Function"

// America_merge
let
    Source = Table.NestedJoin(America,{"Importer_"},MasterSheet,{"Importer_"},"MasterSheet",JoinKind.LeftOuter),
    #"Expanded MasterSheet" = Table.ExpandTableColumn(Source, "MasterSheet", {"Standard Manufacturing CO", "Type", "Remarks"}, {"master.Standard Manufacturing CO", "master.Type", "master.Remarks"})
in
    #"Expanded MasterSheet"

// Britain_merge
let
    Source = Table.NestedJoin(Britain,{"Importer_"},MasterSheet,{"Importer_"},"MasterSheet",JoinKind.LeftOuter),
    #"Expanded MasterSheet" = Table.ExpandTableColumn(Source, "MasterSheet", {"Standard Manufacturing CO", "Type", "Remarks"}, {"master.Standard Manufacturing CO", "master.Type", "master.Remarks"})
in
    #"Expanded MasterSheet"

// Canada_merge
let
    Source = Table.NestedJoin(Canada,{"Importer_"},MasterSheet,{"Importer_"},"MasterSheet",JoinKind.LeftOuter),
    #"Expanded MasterSheet" = Table.ExpandTableColumn(Source, "MasterSheet", {"Standard Manufacturing CO", "Type", "Remarks"}, {"master.Standard Manufacturing CO", "master.Type", "master.Remarks"})
in
    #"Expanded MasterSheet"

// Denmark_merge
let
    Source = Table.NestedJoin(Denmark,{"Importer_"},MasterSheet,{"Importer_"},"MasterSheet",JoinKind.LeftOuter),
    #"Expanded MasterSheet" = Table.ExpandTableColumn(Source, "MasterSheet", {"Standard Manufacturing CO", "Type", "Remarks"}, {"master.Standard Manufacturing CO", "master.Type", "master.Remarks"})
in
    #"Expanded MasterSheet"

// Invoked Function
let
    Source = GetValue("Config",0)
in
    Source

// Invoked Function (2)
let
    Source = GetValue("Config", 1)
in
    Source
