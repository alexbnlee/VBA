Attribute VB_Name = "NewMacros"

Sub ÖÐÓ¢ÎÄÃ°ºÅÇÐ»»()
'
' ÖÐÓ¢ÎÄÃ°ºÅÇÐ»» ºê
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ":"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
