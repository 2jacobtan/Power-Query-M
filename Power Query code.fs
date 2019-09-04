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
            Output2 = Text.Start(Output,12)
        in
            Output2 as text
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
    Source = Excel.Workbook(File.Contents("O:\06 Data Cleaning\" & "masterlist" & ".xlsx"), null, true),
    Sheet1_Sheet = Source{[Item="Sheet1",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Sheet1_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Importer Company", type text}, {"Standard Manufacturing CO", type text}, {"Type", type text}, {"Remarks", type text}, {"Column5", type any}, {"Column6", type any}, {"Column7", type any}, {"Column8", type any}, {"Column9", type any}, {"IND", type text}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Changed Type",{"Importer Company", "Standard Manufacturing CO", "Type", "Remarks"}),
    #"Invoked Custom Function" = Table.AddColumn(#"Removed Other Columns", "Importer_", each #"StripPunctuation(text)"([Importer Company])),
    #"Removed Duplicates" = Table.Distinct(#"Invoked Custom Function", {"Importer_"})
in
    #"Removed Duplicates"

// Raw data files
let
    Source = Folder.Files("O:\05 Newly received data"),
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
