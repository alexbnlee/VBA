Sub 删除全部答案()
    '针对不同的 Highlight 和 颜色 可以替换为白底白字
    Selection.Find.ClearFormatting
    'Selection.Find.Font.Color = wdColorRed
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Color = wdColorWhite
    Selection.Find.Replacement.Highlight = False
    With Selection.Find
        .Text = "：（*）"
        .Replacement.Text = "^&"
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
