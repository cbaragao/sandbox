//-----------------------------------------------------------------------
/* 
Author: Chris Aragao
Title: RemoveSpecialsFromString
Purpose: Remove all special characters from text column
*/

(STRING as text) => let Source = 
if STRING = null then "" else Text.From(STRING),
Replace1 = if Text.Contains(Source, ",") then Text.Replace(Source, ",", "") else Source,
Replace2 = if Text.Contains(Replace1, ".") then Text.Replace(Replace1, ".", "") else Replace1,
Replace3 = if Text.Contains(Replace2, "/") then	Text.Replace(Replace2, "/", "") else Replace2,
Replace4 = if Text.Contains(Replace3, "<") then	Text.Replace(Replace3, "<", "") else Replace3,
Replace5 = if Text.Contains(Replace4, ">") then	Text.Replace(Replace4, ">", "") else Replace4,
Replace6 = if Text.Contains(Replace5, "?") then	Text.Replace(Replace5, "?", "") else Replace5,
Replace7 = if Text.Contains(Replace6, ";") then	Text.Replace(Replace6, ";", "") else Replace6,
Replace8 = if Text.Contains(Replace7, "'") then	Text.Replace(Replace7, "'", "") else Replace7,
Replace9 = if Text.Contains(Replace8, ":") then	Text.Replace(Replace8, ":", "") else Replace8,
Replace10 = if Text.Contains(Replace9, """") then Text.Replace(Replace9, """", "") else Replace9,
Replace11 = if Text.Contains(Replace10, "[") then Text.Replace(Replace10, "[", "") else Replace10,
Replace12 = if Text.Contains(Replace11, "]") then Text.Replace(Replace11, "]", "") else Replace11,
Replace13 = if Text.Contains(Replace12, "{") then Text.Replace(Replace12, "{", "") else Replace12,
Replace14 = if Text.Contains(Replace13, "}") then Text.Replace(Replace13, "}", "") else Replace13,
Replace15 = if Text.Contains(Replace14, "`") then Text.Replace(Replace14, "`", "") else Replace14,
Replace16 = if Text.Contains(Replace15, "~") then Text.Replace(Replace15, "~", "") else Replace15, 
Replace17 = if Text.Contains(Replace16, "!") then Text.Replace(Replace16, "!", "") else Replace16,
Replace18 = if Text.Contains(Replace17, "@") then Text.Replace(Replace17, "@", "") else Replace17,
Replace19 = if Text.Contains(Replace18, "#") then Text.Replace(Replace18, "#", "") else Replace18,
Replace20 = if Text.Contains(Replace19, "$") then Text.Replace(Replace19, "$", "") else Replace19,
Replace21 = if Text.Contains(Replace20, "%") then Text.Replace(Replace20, "%", "") else Replace20,
Replace22 = if Text.Contains(Replace21, "^") then Text.Replace(Replace21, "^", "") else Replace21,
Replace23 = if Text.Contains(Replace22, "&") then Text.Replace(Replace22, "&", "") else Replace22,
Replace24 = if Text.Contains(Replace23, "*") then Text.Replace(Replace23, "*", "") else Replace23,
Replace25 = if Text.Contains(Replace24, "(") then Text.Replace(Replace24, "(", "") else Replace24,
Replace26 = if Text.Contains(Replace25, ")") then Text.Replace(Replace25, ")", "") else Replace25,
Replace27 = if Text.Contains(Replace26, "_") then Text.Replace(Replace26, "_", "") else Replace26,
Replace28 = if Text.Contains(Replace27, "+") then Text.Replace(Replace27, "+", "") else Replace27,
Replace29 = if Text.Contains(Replace28, "|") then Text.Replace(Replace28, "|", "") else Replace28,
Replace30 = if Text.Contains(Replace29, "-") then Text.Replace(Replace29, "-", " ") else Replace29,
Replace31 = if Text.Contains(Replace30, "=") then Text.Replace(Replace30, "=", "") else Replace30,
Replace32 = if Text.Contains(Replace31, "\") then Text.Replace(Replace1, "\", "") else Replace31

in
	Replace32
  
//-----------------------------------------------------------------------
