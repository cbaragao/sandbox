// fnGetRecommendations

(PageStart as text)=>


let
    Source = Web.BrowserContents("https://www.oversight.gov/recommendations?page=" & PageStart),
    #"Extracted Table From Html" = Html.Table(Source, {{"Column1", ".views-field-field-rec-number *"}, {"Column2", ".ui-accordion-content A"}, {"Column3", ".cigie-rec-list-item-row1"}, {"Column4", ".date-display-single"}, {"Column5", ".cigie-rec-list-item-agency *"}, {"Column6", ".cigie-rec-list-item-title"}, {"Column7", ".cigie-rec-list-item-type"}, {"Column8", ".cigie-rec-list-item-type *"}, {"Column9", ".cigie-rec-list-item-report-link *"}, {"Column10", ".cigie-rec-list-item-rec-count"}}, [RowSelector=".views-field-field-rec-number *"]),
    #"Changed Type" = Table.TransformColumnTypes(#"Extracted Table From Html",{{"Column1", type text}, {"Column2", type text}, {"Column3", type text}, {"Column4", type date}, {"Column5", type text}, {"Column6", type text}, {"Column7", type text}, {"Column8", type text}, {"Column9", type text}, {"Column10", type text}})
in
    #"Changed Type"


---------------------------------------------------------------------------------------------

// Results

let
    Source = Web.BrowserContents("https://www.oversight.gov/recommendations"),
    #"Extracted Table From Html" = Html.Table(Source, {{"Results", ".oversight-table-summary"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Extracted Table From Html",{{"Results", type text}}),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Changed Type", "Results", Splitter.SplitTextByEachDelimiter({"of "}, QuoteStyle.Csv, false), {"Results.1", "Results.2"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Split Column by Delimiter",{{"Results.1", type text}, {"Results.2", type text}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type1",{"Results.1"}),
    #"Split Column by Delimiter1" = Table.SplitColumn(#"Removed Columns", "Results.2", Splitter.SplitTextByEachDelimiter({" "}, QuoteStyle.Csv, false), {"Results.2.1", "Results.2.2"}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Split Column by Delimiter1",{{"Results.2.1", Int64.Type}, {"Results.2.2", type text}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Changed Type2",{"Results.2.2"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns1",{{"Results.2.1", "Results"}}),
    #"Added Custom" = Table.AddColumn(#"Renamed Columns", "ResultsDivBy20", each [Results]/20),
    #"Inserted Round Up" = Table.AddColumn(#"Added Custom", "Round Up", each Number.RoundUp([ResultsDivBy20]), Int64.Type),
    #"Removed Columns2" = Table.RemoveColumns(#"Inserted Round Up",{"Results", "ResultsDivBy20"}),
    #"Renamed Columns1" = Table.RenameColumns(#"Removed Columns2",{{"Round Up", "Results"}})
in
    #"Renamed Columns1"

-----------------------------------------------------------------------------------------------------

// OversightRecommendations

let
    Source = {0..(Table.FirstValue(Results)-1)},
    #"Converted to Table" = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Renamed Columns" = Table.RenameColumns(#"Converted to Table",{{"Column1", "Index"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns",{{"Index", type text}}),
    #"Invoked Custom Function" = Table.AddColumn(#"Changed Type", "Recommendations", each fxGetRecommendations([Index])),
    #"Expanded Recommendations" = Table.ExpandTableColumn(#"Invoked Custom Function", "Recommendations", {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10"}, {"Recommendations.Column1", "Recommendations.Column2", "Recommendations.Column3", "Recommendations.Column4", "Recommendations.Column5", "Recommendations.Column6", "Recommendations.Column7", "Recommendations.Column8", "Recommendations.Column9", "Recommendations.Column10"}),
    #"Renamed Columns1" = Table.RenameColumns(#"Expanded Recommendations",{{"Index", "Page"}}),
    #"Added Index" = Table.AddIndexColumn(#"Renamed Columns1", "Index", 1, 1, Int64.Type),
    #"Reordered Columns" = Table.ReorderColumns(#"Added Index",{"Index", "Page", "Recommendations.Column1", "Recommendations.Column2", "Recommendations.Column3", "Recommendations.Column4", "Recommendations.Column5", "Recommendations.Column6", "Recommendations.Column7", "Recommendations.Column8", "Recommendations.Column9", "Recommendations.Column10"}),
    #"Renamed Columns2" = Table.RenameColumns(#"Reordered Columns",{{"Recommendations.Column1", "AgencyRecID"}, {"Recommendations.Column2", "Recommendation"}, {"Recommendations.Column3", "Date_Agency"}, {"Recommendations.Column4", "Date"}, {"Recommendations.Column5", "Agency"}, {"Recommendations.Column6", "ReportTitle"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns2",{"Recommendations.Column7"}),
    #"Renamed Columns3" = Table.RenameColumns(#"Removed Columns",{{"Recommendations.Column8", "OversightType"}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Renamed Columns3",{"Recommendations.Column9"}),
    #"Renamed Columns4" = Table.RenameColumns(#"Removed Columns1",{{"Recommendations.Column10", "Open?"}}),
    #"Added Custom" = Table.AddColumn(#"Renamed Columns4", "Source", each "Oversight.gov"),
    #"Renamed Columns5" = Table.RenameColumns(#"Added Custom",{{"Agency", "Reporting Agency"}}),
    #"Removed Columns2" = Table.RemoveColumns(#"Renamed Columns5",{"AgencyRecID", "Date_Agency", "OversightType", "Open?"}),
    #"Filled Up" = Table.FillUp(#"Removed Columns2",{"Index", "Date", "Reporting Agency", "ReportTitle"}),
    #"Duplicated Column" = Table.DuplicateColumn(#"Filled Up", "Reporting Agency", "Reporting Agency - Copy"),
    #"Renamed Columns6" = Table.RenameColumns(#"Duplicated Column",{{"Reporting Agency - Copy", "Agency"}}),
    #"Added Custom1" = Table.AddColumn(#"Renamed Columns6", "Status", each "Open")
in
    #"Added Custom1"


----------------------------------------------------------------------------------------------------

// GAORecommendations

let
    Source = Csv.Document(Web.Contents("https://www.gao.gov/open-recs-csv?processed=1&topic=all&page&rows=10000"),[Delimiter=",", Columns=10, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Column1", type text}, {"Column2", type text}, {"Column3", type text}, {"Column4", type text}, {"Column5", type text}, {"Column6", type text}, {"Column7", type text}, {"Column8", type text}, {"Column9", type text}, {"Column10", type text}}),
    #"Removed Top Rows" = Table.Skip(#"Changed Type",5),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars=true]),
    #"Changed Type1" = Table.TransformColumnTypes(#"Promoted Headers",{{"Publication Name", type text}, {"Publication  Number", type text}, {"Date Publication Issued", type date}, {"Director Name", type text}, {"Director Phone", type text}, {"Agency", type text}, {"Recommendation", type text}, {"Status", type text}, {"Priority", type text}, {"Comments", type text}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type1",{{"Publication Name", "ReportTitle"}, {"Publication  Number", "ReportNumber"}, {"Date Publication Issued", "Date"}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Renamed Columns",{"ReportTitle", "ReportNumber", "Date", "Agency", "Recommendation", "Status"}),
    #"Filtered Rows" = Table.SelectRows(#"Removed Other Columns", each ([ReportNumber] <> "")),
    #"Added Index" = Table.AddIndexColumn(#"Filtered Rows", "Index", 1, 1, Int64.Type),
    #"Reordered Columns" = Table.ReorderColumns(#"Added Index",{"Index", "ReportTitle", "ReportNumber", "Date", "Agency", "Recommendation", "Status"}),
    #"Added Custom" = Table.AddColumn(#"Reordered Columns", "Source", each "GAO")
in
    #"Added Custom"
----------------------------------------------------------------------------------------------------------------

// GlobalRecommendations

let
    Source = Table.Combine({OversightRecommendations, GAORecommendations}),
    #"Cleaned Text" = Table.TransformColumns(Source,{{"Recommendation", Text.Clean, type text}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Cleaned Text",{{"Date", type date}})
in
    #"Changed Type"

