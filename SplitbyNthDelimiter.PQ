//-----------------------------------------------------------------------
/* 
Author: @Jimmy801 in Microsoft Power BI Community Forum
Title: Split by Nth Delimiter
Purpose: Split text by nth delimiter in the text
Found at: https://community.powerbi.com/t5/user/viewprofilepage/user-id/160408
*/




(tSplitText as text, tDelimiter as text, nOccurance as nullable number) as list =>
    let
        nOccuranceIntern = if nOccurance = null then 2 else nOccurance,
        Source =  if tSplitText = null then "" else Text.From(tSplitText),
        ListFromText = Text.ToList(tSplitText),
        CreateListOfDelimiter = List.Accumulate
        (
            ListFromText,
            [CurrentRow = -1, DelimiterFound={}],
            (state, current)=> 
                if current = tDelimiter then
                    Record.TransformFields(state, {{"DelimiterFound", each _ & {state[CurrentRow]}},{"CurrentRow", each _ +1 }})
                else
                    Record.TransformFields(state, {{"CurrentRow", each _ +1 }})
        )[DelimiterFound],
        ListAlternate = List.Alternate(CreateListOfDelimiter,nOccuranceIntern-1,nOccuranceIntern-1),
        ChangeTextList = List.Accumulate
        (
            ListAlternate,
            ListFromText,
            (state, current)=>
            List.ReplaceRange(state,current+1,1,{"%%%"})
        ),
        
        SplitTextByNewDelimiter = Text.Split
        (
            List.Accumulate
                (
                    ChangeTextList,
                    "",
                    (state, current)=>
                    state & current
                ), 
            "%%%"
         ) 

in 
        SplitTextByNewDelimiter

//-----------------------------------------------------------------------
