Sub Yellow()

    '选定光标所在行
    Selection.Expand Unit:=wdParagraph
    '选定行背景色设置
    Selection.Range.HighlightColorIndex = wdYellow
    '选定行字体颜色设置
    Selection.Range.Font.ColorIndex = wdRed

End Sub
