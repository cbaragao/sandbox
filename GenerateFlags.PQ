//-----------------------------------------------------------------------
/* 
Author: Chris Aragao
Title: Generate Flags
Purpose: Generate a column of flags based on keywords in current text column
*/

(STRING as text)=> 

// Pull in Source
let Source = if STRING = null then "blank" else Text.From(STRING),

// Convert to lower to make matching easier
LowerSource = Text.Lower(Source),

// Remove special characters
NoSpecials = RemoveSpecialsFromString(LowerSource),

// Remove numbers
NoNums = RemoveNumbersFromString(NoSpecials),

// Remove extra spaces
NoExtraSpaces = RemoveExtraSpacesFromString(NoNums),

// Split the text into a list
SplitList = Text.Split(NoExtraSpaces, " "),

// Take null items out of the list
NoNulls = List.RemoveNulls(SplitList),

// Get a word count for each row
WordCount = GetWordCount(NoExtraSpaces),

// See if the list for each row contains duplicate keywords
MatchDup = if List.ContainsAny(NoNulls, {"duplicate", "duplicated"}) then {"Duplicate"} else {""},

// See if the list for each row contains ROI keywords
MatchROI = if List.Contains(NoNulls, {"roi"}) then List.Combine({MatchDup, {"Release of Information"}}) else MatchDup,

// See if the list for each row contains RFI keywords
MatchRFI = if List.Contains(NoNulls, {"rfi"}) then List.Combine({MatchROI, {"Request for Information"}}) else MatchROI,

// Combine the list back up
CombineText = Text.Combine(MatchRFI, "; "),

// See if status check is mentioned
MatchStatus = if Text.Contains(LowerSource, "status check") then (CombineText & "; Status Check") else CombineText,

// Flag blank synopsis
MatchBlank = if LowerSource = "blank" then (MatchStatus & "; Blank Synopsis") else MatchStatus,

// Flag if there is a low word count
LowWordCount = if WordCount <= 20 then (MatchBlank & "; Short Synopsis") else MatchBlank,

// Post-cleaning
FinalText = if Text.StartsWith(LowWordCount, "; ") then Text.RemoveRange(LowWordCount, 0, 2) else LowWordCount,

// Trim the result
FinalTextTrimmed = Text.Trim(FinalText),

// If there are no flags, put in none
Blanks = if FinalTextTrimmed = "" then "None" else FinalTextTrimmed

in
    Blanks

//-----------------------------------------------------------------------
