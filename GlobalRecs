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
    #"Removed Columns2" = Table.RemoveColumns(#"Renamed Columns5",{"AgencyRecID", "Date_Agency", "OversightType", "Open?"})
in
    #"Removed Columns2"

----------------------------------------------------------------------------------------------

// fnGAORecommendations

(PgStart as text)=>


let
    Source = Web.BrowserContents("https://www.gao.gov/reports-testimonies/recommendations-database?processed=1&topic=all&page" & PgStart),
    #"Extracted Table From Html" = Html.Table(Source, {{"Column1", ".views-field-field-agency"}, {"Column2", ".views-field-field-recommendation"}, {"Column3", ".views-field-field-status-code"}, {"Column4", ".code"}, {"Column5", ".comment *"}, {"Column6", "H3 *"}, {"Column7", ".field\-\-name-field-product-number"}, {"Column8", ".datetime"}, {"Column9", ".field\-\-name-field-staff-contact"}, {"Column10", ".recommendations-counter"}, {"Column11", ".badge"}, {"Column12", ".js-toggle"}}, [RowSelector=".views-field-field-agency"]),
    #"Promoted Headers" = Table.PromoteHeaders(#"Extracted Table From Html", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Agency", type text}, {"Recommendation", type text}, {"Status", type text}, {"Column4", type text}, {"Column5", type text}, {"Column6", type text}, {"Column7", type text}, {"Column8", type date}, {"Column9", type text}, {"Column10", type text}, {"Column11", Int64.Type}, {"Column12", type text}})
in
    #"Changed Type"

------------------------------------------------------------------------------------------------

// GAOResults

let
    Source = Web.BrowserContents("https://www.gao.gov/reports-testimonies/recommendations-database?processed=1&topic=all&page=1"),
    #"Extracted Table From Html" = Html.Table(Source, {{"GAOResults", ".result-count"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Extracted Table From Html",{{"GAOResults", type text}}),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Changed Type", "GAOResults", Splitter.SplitTextByEachDelimiter({" of "}, QuoteStyle.Csv, false), {"GAOResults.1", "GAOResults.2"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Split Column by Delimiter",{{"GAOResults.1", type text}, {"GAOResults.2", type text}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type1",{"GAOResults.1"}),
    #"Split Column by Delimiter1" = Table.SplitColumn(#"Removed Columns", "GAOResults.2", Splitter.SplitTextByEachDelimiter({" "}, QuoteStyle.Csv, false), {"GAOResults.2.1", "GAOResults.2.2"}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Split Column by Delimiter1",{{"GAOResults.2.1", Int64.Type}, {"GAOResults.2.2", type text}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Changed Type2",{"GAOResults.2.2"}),
    #"Added Custom" = Table.AddColumn(#"Removed Columns1", "Pages", each [GAOResults.2.1]/20),
    #"Removed Columns2" = Table.RemoveColumns(#"Added Custom",{"GAOResults.2.1"}),
    #"Rounded Up" = Table.TransformColumns(#"Removed Columns2",{{"Pages", Number.RoundUp, Int64.Type}})
in
    #"Rounded Up"

----------------------------------------------------------------------------------------------------

// GAORecommendations

let
    Source = {0..(Table.FirstValue(GAOResults)-1)},
    #"Converted to Table" = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Renamed Columns" = Table.RenameColumns(#"Converted to Table",{{"Column1", "Page"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns",{{"Page", type text}}),
    #"Invoked Custom Function" = Table.AddColumn(#"Changed Type", "Recommendations", each fxGAORecommendations([Page])),
    #"Expanded Recommendations" = Table.ExpandTableColumn(#"Invoked Custom Function", "Recommendations", {"Agency", "Recommendation", "Status", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12"}, {"Recommendations.Agency", "Recommendations.Recommendation", "Recommendations.Status", "Recommendations.Column4", "Recommendations.Column5", "Recommendations.Column6", "Recommendations.Column7", "Recommendations.Column8", "Recommendations.Column9", "Recommendations.Column10", "Recommendations.Column11", "Recommendations.Column12"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Recommendations",{"Recommendations.Column11", "Recommendations.Column12", "Recommendations.Column10"}),
    #"Renamed Columns1" = Table.RenameColumns(#"Removed Columns",{{"Recommendations.Column9", "Contact"}, {"Recommendations.Column8", "Date"}, {"Recommendations.Column7", "ReportNumber"}, {"Recommendations.Column6", "ReportTitle"}}),
    #"Filtered Rows" = Table.SelectRows(#"Renamed Columns1", each ([Recommendations.Status] = "Open#(lf)  #(lf)#(lf)#(lf)When we confirm what actions the agency has taken in response to this recommendation, we will provide updated information.")),
    #"Removed Columns1" = Table.RemoveColumns(#"Filtered Rows",{"Recommendations.Status", "Recommendations.Column5"}),
    #"Renamed Columns2" = Table.RenameColumns(#"Removed Columns1",{{"Recommendations.Column4", "Status"}, {"Recommendations.Recommendation", "Recommendation"}, {"Recommendations.Agency", "Agency"}}),
    #"Added Index" = Table.AddIndexColumn(#"Renamed Columns2", "Index", 1, 1, Int64.Type),
    #"Reordered Columns" = Table.ReorderColumns(#"Added Index",{"Index", "Page", "Agency", "Recommendation", "Status", "ReportTitle", "ReportNumber", "Date", "Contact"}),
    #"Added Custom" = Table.AddColumn(#"Reordered Columns", "Reporting Agency", each "GAO"),
    #"Added Custom1" = Table.AddColumn(#"Added Custom", "Source", each "GAO.gov"),
    #"Removed Columns2" = Table.RemoveColumns(#"Added Custom1",{"Contact", "Status"})
in
    #"Removed Columns2"

----------------------------------------------------------------------------------------------------------------

// GlobalRecommendations

let
    Source = Table.Combine({OversightRecommendations, GAORecommendations}),
    #"Cleaned Text" = Table.TransformColumns(Source,{{"Recommendation", Text.Clean, type text}})
in
    #"Cleaned Text"

