//-----------------------------------------------------------------------
/* 
Author: Chris Aragao
Title: RemoveNumbersFromString
Purpose: Remove all numbers from text column
*/

(STRING as text) => let Source = 
if STRING = null then "" else Text.From(STRING), 
RemoveZero = Text.Replace(Source, "0", ""),
RemoveOne = Text.Replace(RemoveZero, "1", ""),
RemoveTwo = Text.Replace(RemoveOne, "2", ""),
RemoveThree = Text.Replace(RemoveTwo, "3", ""),
RemoveFour = Text.Replace(RemoveThree, "4", ""),
RemoveFive = Text.Replace(RemoveFour, "5", ""),
RemoveSix = Text.Replace(RemoveFive, "6", ""),
RemoveSeven = Text.Replace(RemoveSix, "7", ""),
RemoveEight = Text.Replace(RemoveSeven, "8", ""),
RemoveNine = Text.Replace(RemoveEight, "9", "")

in 
	RemoveNine

//-----------------------------------------------------------------------
