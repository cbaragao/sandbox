// Get SharePointVersionHistory - Power Query

// Function/Query:
Let Source = (
    VersionsRelevantSharePointListName as text, 
    VersionsRelevantSharePointLocation as text, 
    VersionsRelevantItemID as number
) => let Source = 

Xml.Tables(
  Web.Contents(
    Text.Combine(
      {
        VersionsRelevantSharePointLocation,"/_api/web/Lists/getbytitle(
          '",VersionsRelevantSharePointListName ,"'
        )/items(
          ",Text.From(
            VersionsRelevantItemID
           ),"
        )/versions"
       }
     )
    )
  ),
  
  entry = Source{0}[entry],
  #"Removed Other Columns2" = Table.SelectColumns(entry,{"content"}),
  #"Expanded content" = Table.ExpandTableColumn(#"Removed Other Columns2", "content", {"http://schemas.microsoft.com/ado/2007/08/dataservices/metadata"}, {"content"}),
  #"Expanded content1" = Table.ExpandTableColumn(#"Expanded content", "content", {"properties"}, {"properties"}),
  #"Expanded properties" = Table.ExpandTableColumn(#"Expanded content1", "properties", {"http://schemas.microsoft.com/ado/2007/08/dataservices"}, {"properties"}) 
  
  in 
    
    #"Expanded properties" 
    
  in 
    Source

// From <https://community.powerbi.com/t5/Desktop/Getting-SharePoint-List-items-with-full-history-version/td-p/64743> 
