//-----------------------------------------------------------------------
/* 
Author: Chris Aragao
Title: GetWordCount
Purpose: Get a word count of words in a list
*/

// Get source
(STRING as text)=> 
let Source = if STRING = "" then "" else Text.From(STRING),

// Remove special characters
NoSpecials = RemoveSpecialsFromString(Source),

// Take out numbers
NoNums = RemoveNumbersFromString(NoSpecials),

// Split to list for counting
SplitList = Text.Split(NoNums, " "),

// Take out null items if any
NoNulls = List.RemoveNulls(SplitList),

// Count the list
ListCount = List.Count(NoNulls)

in 
    ListCount

//-----------------------------------------------------------------------
