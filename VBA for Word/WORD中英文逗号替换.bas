Attribute VB_Name = "NewMacros"
Sub ÖÐÓ¢ÎÄ¶ººÅÇÐ»»()
Attribute ÖÐÓ¢ÎÄ¶ººÅÇÐ»».VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.ÖÐÓ¢ÎÄ¶ººÅÇÐ»»"
'
' ÖÐÓ¢ÎÄ¶ººÅÇÐ»» ºê
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ","
        .Replacement.Text = "£¬"
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
