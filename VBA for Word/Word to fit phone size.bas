Sub PHONE()

    Selection.WholeStory
    With ActiveDocument.Styles(wdStyleNormal).Font
        If .NameFarEast = .NameAscii Then
            .NameAscii = ""
        End If
        .NameFarEast = ""
    End With
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        .TopMargin = CentimetersToPoints(0.51)
        .BottomMargin = CentimetersToPoints(0.52)
        .LeftMargin = CentimetersToPoints(0.51)
        .RightMargin = CentimetersToPoints(0.51)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(1.5)
        .FooterDistance = CentimetersToPoints(1.75)
        .PageWidth = CentimetersToPoints(21)
        .PageHeight = CentimetersToPoints(29.7)
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = False
        .TwoPagesOnOne = False
        .BookFoldPrinting = False
        .BookFoldRevPrinting = False
        .BookFoldPrintingSheets = 1
        .GutterPos = wdGutterPosLeft
        .LayoutMode = wdLayoutModeLineGrid
    End With
    CommandBars("Navigation").Visible = False
    Selection.WholeStory
    Selection.Font.Size = 15
    Selection.Font.Size = 16
    Selection.Font.Grow
    Selection.Font.Grow
    Selection.Font.Grow
    Selection.Font.Grow

    '文件另存为指定目录
    Dim filename As String
    filename = ActiveDocument.Name

    ChangeFileOpenDirectory _
        "D:\01-Working\YifangCloud\FangCloudSync\个人文件\吉林-考试\PHONE\"
    ActiveDocument.SaveAs2 filename:=filename, FileFormat:= _
        wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, CompatibilityMode:=15
    
End Sub
