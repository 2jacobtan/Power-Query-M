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
    Source = Excel.Workbook(File.Contents("O:\Docs\00.Business Planning\Export Genius\Data\06 Data Cleaning\" & "masterlist USA" & ".xlsx"), null, true),
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
    Source = Folder.Files("O:\Docs\00.Business Planning\Export Genius\Data\05 Newly received data"),
    #"Filtered Hidden Files1" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true),
    #"Filtered Rows" = Table.SelectRows(#"Filtered Hidden Files1", each Text.StartsWith([Extension], ".xls")),
    #"Invoked Custom Function" = Table.AddColumn(#"Filtered Rows", "workbookSheet1(binary)", each #"workbookSheet1(binary)"([Content])),
    #"Removed Other Columns" = Table.SelectColumns(#"Invoked Custom Function",{"Name", "workbookSheet1(binary)"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Other Columns",{{"Name", "Source.Name"}})
in
    #"Renamed Columns"

// India
let
    Source = #"Raw data files",
    #"CountryFilterExpand" = CountryFilterExpand(Source,"India"),
    #"Invoked Custom Function" = Table.AddColumn(CountryFilterExpand, "Importer_", each #"StripPunctuation(text)"([FOREIGN IMPORTER NAME]))
in
    #"Invoked Custom Function"

// Indonesia
let
    Source = #"Raw data files",
    #"CountryFilterExpand" = CountryFilterExpand(Source,"Indonesia"),
    #"Invoked Custom Function" = Table.AddColumn(CountryFilterExpand, "Importer_", each #"StripPunctuation(text)"([Importer]))
in
    #"Invoked Custom Function"

// Pakistan
let
    Source = #"Raw data files",
    #"CountryFilterExpand" = CountryFilterExpand(Source,"Pakistan"),
    #"Invoked Custom Function" = Table.AddColumn(CountryFilterExpand, "Importer_", each #"StripPunctuation(text)"([IMPORTER NAME]))
in
    #"Invoked Custom Function"

// Vietnam
let
    Source = #"Raw data files",
    #"CountryFilterExpand" = CountryFilterExpand(Source,"Vietnam"),
    #"Invoked Custom Function" = Table.AddColumn(CountryFilterExpand, "Importer_", each #"StripPunctuation(text)"([Importer]))
in
    #"Invoked Custom Function"

// India_merge
let
    Source = Table.NestedJoin(India,{"Importer_"},MasterSheet,{"Importer_"},"MasterSheet",JoinKind.LeftOuter),
    #"Expanded MasterSheet" = Table.ExpandTableColumn(Source, "MasterSheet", {"Standard Manufacturing CO", "Type", "Remarks"}, {"master.Standard Manufacturing CO", "master.Type", "master.Remarks"})
in
    #"Expanded MasterSheet"

// Indonesia_merge
let
    Source = Table.NestedJoin(Indonesia,{"Importer_"},MasterSheet,{"Importer_"},"MasterSheet",JoinKind.LeftOuter),
    #"Expanded MasterSheet" = Table.ExpandTableColumn(Source, "MasterSheet", {"Standard Manufacturing CO", "Type", "Remarks"}, {"master.Standard Manufacturing CO", "master.Type", "master.Remarks"})
in
    #"Expanded MasterSheet"

// Pakistan_merge
let
    Source = Table.NestedJoin(Pakistan,{"Importer_"},MasterSheet,{"Importer_"},"MasterSheet",JoinKind.LeftOuter),
    #"Expanded MasterSheet" = Table.ExpandTableColumn(Source, "MasterSheet", {"Standard Manufacturing CO", "Type", "Remarks"}, {"master.Standard Manufacturing CO", "master.Type", "master.Remarks"})
in
    #"Expanded MasterSheet"

// Vietnam_merge
let
    Source = Table.NestedJoin(Vietnam,{"Importer_"},MasterSheet,{"Importer_"},"MasterSheet",JoinKind.LeftOuter),
    #"Expanded MasterSheet" = Table.ExpandTableColumn(Source, "MasterSheet", {"Standard Manufacturing CO", "Type", "Remarks"}, {"master.Standard Manufacturing CO", "master.Type", "master.Remarks"})
in
    #"Expanded MasterSheet"
