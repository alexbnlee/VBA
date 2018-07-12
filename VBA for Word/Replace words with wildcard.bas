Sub Replace_words()

    Selection.Find.ClearFormatting
    Selection.Find.Font.Color = wdColorRed
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Color = wdColorYellow
    With Selection.Find
        .text = "（*）.您选择了：（*）"
        .Replacement.text = "^&"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

End Sub
