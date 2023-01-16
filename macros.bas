Attribute VB_Name = "NewMacros"
Option Explicit

Sub CreateWikiLink()

  i = "Herzschwaeche"
  p = "ae"
  o = Left$(i, InStr(1, i, p) - 1) + "ä" + Right$(i, Len(i) - InStr(1, i, p) - 1)
  Debug.Print o
  

End Sub
Sub CreateTOC()

  Dim sHeader, WikiLinks() As String
  Dim i
  Dim p As Object
  
  Open "c:\tmp\toc.txt" For Output As #1
  
  Selection.HomeKey Unit:=wdStory
  Set p = ActiveDocument.Paragraphs
  ReDim WikiLinks(p.Count())
  
  For i = 1 To p.Count()
    p(i).Range.Select
    sHeader = Selection.Text
    sHeader = Left$(sHeader, Len(sHeader) - 1)
    If p(i).Style = "Verzeichnis 3" Then
      WikiLinks(i) = "[yadawiki link=" + Chr$(34) + sHeader + Chr$(34) + " show=" + Chr$(34) + sHeader + Chr$(34) + "]"
    Else
      WikiLinks(i) = sHeader
    End If
    Print #1, WikiLinks(i)
  Next i
  
  Close #1

End Sub
Sub CheckLinks()
'
' CheckLinks Makro
'
'
  Dim i As Integer
  Dim sLinkText As String
  Dim iNumber As Integer
  
  iNumber = 4520
  
  For i = 1 To 4552
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Link")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    sLinkText = Selection.Text
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    
    If Selection.Text <> "%" Or Selection.Style <> "Verborgen" Then
      MsgBox "Inconsistent Formatting at: " + sLinkText
      End
    End If
    
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    
  Next i

End Sub
Sub CreateLinkIndex()
'
' CreateLinkIndex Makro
'
'
  Dim i As Integer
  Dim sLinkText As String
  
  Open "c:\tmp\index.txt" For Output As #1
  
  For i = 1 To 4544
  
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Verborgen")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    sLinkText = Selection.Text
    sLinkText = Left$(Right$(Selection.Text, Len(Selection.Text) - 1), Len(Selection.Text) - 6)
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    
    Print #1, sLinkText

  Next i
  
  Close #1
  
End Sub
Sub CreateYadaWikiLink()
'
' CreateYadaWikiLink Makro
'
'
  Dim i As Integer
  Dim sLink As String
  Dim sShow As String
  Dim b As Boolean
  
  b = DoesStyleExist("YadaWikiLink", ActiveDocument)
  
  If b = False Then
    CreateWikiStylesheet
  End If
    
  For i = 1 To 4424
    
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Link")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    sShow = Selection.Text
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Verborgen")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    sLink = Left$(Right$(Selection.Text, Len(Selection.Text) - 1), Len(Selection.Text) - 6)
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.Style = ActiveDocument.Styles("YadaWikiLink")
    Selection.TypeText Text:="[yadawiki link=" + Chr$(34) + sLink + Chr$(34) + " show=" + Chr$(34) + sShow + Chr$(34) + "]"
  
  Next i

End Sub

Sub DeleteUnusedStyles()
    Dim oStyle As Style

    For Each oStyle In ActiveDocument.Styles
        'Only check out non-built-in styles
        If oStyle.BuiltIn = False Then
            With ActiveDocument.Content.Find
                .ClearFormatting
                .Style = oStyle.NameLocal
                .Execute FindText:="", Format:=True
                If .Found = False Then oStyle.Delete
            End With
        End If
    Next oStyle
End Sub

Sub DeleteAllBookmarks()
    Dim objBookmark As Bookmark

    For Each objBookmark In ActiveDocument.Bookmarks
        objBookmark.Delete
    Next
End Sub

Sub RenameImagesInDocument()

  Dim i As Integer
  Dim sLine As String
  Dim iPos As Integer
  Dim sReplacementText As String
  
  Const NUMBER_OF_IMAGES = 579
  
  Dim old_name(NUMBER_OF_IMAGES) As String
  Dim new_name(NUMBER_OF_IMAGES) As String
  
  Open "d:\projekte\galeria anatomica\website\picture_rename_list.csv" For Input As #1
  
  For i = 1 To NUMBER_OF_IMAGES
    Line Input #1, sLine
    iPos = InStr(1, sLine, ";")
    old_name(i) = Left$(sLine, iPos - 1)
    new_name(i) = Right$(sLine, Len(sLine) - iPos)
  Next i
  
  Close #1

  For i = 1 To NUMBER_OF_IMAGES
  
    sReplacementText = "<img src=#!!!#http://192.168.178.65:9980/wp-content/uploads/" + new_name(i) + "#!!!#"
    
    Debug.Print "Replacing item " + Str$(i) + " of " + Trim(Str$(NUMBER_OF_IMAGES))
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("GrafikEingebunden")
    Selection.Find.Replacement.Style = ActiveDocument.Styles("GrafikEingebunden")
    With Selection.Find
        .Text = "{" + old_name(i) + "}"
        .Replacement.Text = sReplacementText
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
  
  Next i
  
  Debug.Print "Search-Replace process finished."

' <img src="http://192.168.178.65:9980/wp-content/uploads/01-01-000-04-L4C-01-768x542.jpg" alt="Bau und Funktion des Kehlkopfes" />

End Sub


Sub RenameImagesInDocument_2()

  Dim i As Integer
  Dim sLine As String
  Dim iPos As Integer
  Dim sReplacementText As String
  Dim LineItems() As String
  
  Const NUMBER_OF_IMAGES = 579
  
  Dim old_name(NUMBER_OF_IMAGES) As String
  Dim title(NUMBER_OF_IMAGES) As String
  Dim new_name(NUMBER_OF_IMAGES) As String
  
  Open "d:\projekte\galeria anatomica\website\pictures_labeled_579.CSV" For Input As #1
  
  For i = 1 To NUMBER_OF_IMAGES
    Line Input #1, sLine
    LineItems = Split(sLine, ";")
    old_name(i) = LineItems(0)
    title(i) = LineItems(1)
    new_name(i) = LineItems(2)
  Next i
  
  Close #1

  For i = 1 To NUMBER_OF_IMAGES
  
    sReplacementText = "<img src=" + Chr$(34) + "http://192.168.178.65:9980/wp-content/uploads/" + new_name(i) + Chr$(34) + " alt=" + Chr$(34) + title(i) + Chr$(34) + " />"
    
    Debug.Print "Replacing item " + Str$(i) + " of " + Trim(Str$(NUMBER_OF_IMAGES))
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("GrafikEingebunden")
    Selection.Find.Replacement.Style = ActiveDocument.Styles("GrafikEingebunden")
    With Selection.Find
        .Text = "{" + old_name(i) + "}"
        .Replacement.Text = sReplacementText
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
  
  Next i
  
  Debug.Print "Search-Replace process finished."

' <img src="http://192.168.178.65:9980/wp-content/uploads/01-01-000-04-L4C-01-768x542.jpg" alt="Bau und Funktion des Kehlkopfes" />

End Sub

Sub CreateWikiStylesheet()
Attribute CreateWikiStylesheet.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.CreateWikiStylesheet"
'
' CreateWikiStylesheet Makro
'
'
  ActiveDocument.Styles.Add Name:="YadaWikiLink", Type:=wdStyleTypeCharacter
  With ActiveDocument.Styles("YadaWikiLink").Font
    .Name = "Calibri"
    .Color = -738131969
  End With
  With ActiveDocument.Styles("YadaWikiLink").Font
    With .Shading
      .Texture = wdTextureNone
      .ForegroundPatternColor = wdColorAutomatic
      .BackgroundPatternColor = wdColorAutomatic
    End With
    .Borders(1).LineStyle = wdLineStyleNone
    .Borders.Shadow = False
  End With
    
End Sub

Function DoesStyleExist(ByVal styleToTest As String, ByVal docToTest As Word.Document) As Boolean

  Dim testStyle As Word.Style
    
  On Error Resume Next
  Set testStyle = docToTest.Styles(styleToTest)
  DoesStyleExist = Not testStyle Is Nothing

End Function
Sub MoveChapterInfo()
Attribute MoveChapterInfo.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.MoveChapterInfo"
'
' MoveChapterInfo Makro
'
'
  Dim i As Integer
  
  For i = 1 To 1708
    
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Kapitelinfo")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.Cut
    Selection.MoveDown Unit:=wdParagraph, Count:=1
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    
  Next i
    
End Sub
Sub UnhideLink()
Attribute UnhideLink.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.UnhideLink"
'
' UnhideLink Makro
'
'
  Dim i As Integer
  
  For i = 1 To 4418
    
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("YadaWikiLink")
    With Selection.Find
        .Text = "show="""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("YadaWikiLink")
    Selection.ExtendMode = True
    With Selection.Find
        .Text = """]"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
    With Selection.Font
        .Name = "Calibri Light"
        .Size = 12
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = -738131969
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Spacing = 0
        .Scaling = 100
        .Position = 0
        .Kerning = 0
        .Animation = wdAnimationNone
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    
    Selection.ExtendMode = False
    Selection.MoveRight Unit:=wdCharacter, Count:=3

  Next i

End Sub
Sub CreateHeaderBookmark()
Attribute CreateHeaderBookmark.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.CreateHeaderBookmark"
'
' CreateHeaderBookmark Makro
'
'
  Dim i As Integer, sBookmark As String
    
  For i = 1 To 14
        
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Überschrift 3")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    sBookmark = Selection.Text
    sBookmark = Replace(sBookmark, "_", "")
    sBookmark = Replace(sBookmark, " ", "_")
    sBookmark = Replace(sBookmark, "-", "_")
    sBookmark = Replace(sBookmark, ".", "_")
    sBookmark = Replace(sBookmark, "(", "_")
    sBookmark = Replace(sBookmark, ")", "_")
    sBookmark = Replace(sBookmark, ",", "_")
    sBookmark = Replace(sBookmark, "?", "_")
    sBookmark = Replace(sBookmark, "/", "_")
    sBookmark = Replace(sBookmark, "!", "_")
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:=sBookmark
        .DefaultSorting = wdSortByName
        .ShowHidden = False
    End With
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    
  Next i
End Sub
Sub CleanseChapterInfo()
Attribute CleanseChapterInfo.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.CleanseChapterInfo"
'
' CleanseChapterInfo Makro
'
'
  Dim i As Integer
  
  For i = 1 To 1698
      
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Kapitelinfo")
    With Selection.Find
        .Text = "Level^t"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
  ' Selection.HomeKey Unit:=wdLine
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.ExtendMode = True
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Kapitelinfo")
    With Selection.Find
        .Text = "^l"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.ExtendMode = False
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:="STYLEREF  ""Überschrift 1"" ", PreserveFormatting:=True
    
  Next i

End Sub

Sub SplitDocument()
Attribute SplitDocument.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.SplitDocument"
'
' SplitDocument Makro
'
'
  Dim i As Integer
  Dim sDocumentName As String
  Dim sHeadingNumber As String
  
  Selection.HomeKey Unit:=wdStory
  
  For i = 1 To 117
    
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Überschrift 2")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    
    sDocumentName = ""
    sHeadingNumber = ""
    sDocumentName = sDocumentName + Left$(Selection.Text, Len(Selection.Text) - 1)
    sDocumentName = Replace(sDocumentName, "Der ", "")
    sDocumentName = Replace(sDocumentName, "Die ", "")
    sDocumentName = Replace(sDocumentName, "Das ", "")
    sHeadingNumber = Selection.Range.ListFormat.ListString
    Selection.MoveDown Unit:=wdParagraph, Count:=1
    Selection.ExtendMode = True
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Überschrift 2")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.MoveUp Unit:=wdParagraph, Count:=1
    Selection.Copy
    Documents.Add Template:="GaleriaAnatomicaWiki", NewTemplate:=False, DocumentType:=0
  ' Documents.Add NewTemplate:=False, DocumentType:=wdNewBlankDocument
    Selection.PasteAndFormat (wdPasteDefault)
    Selection.TypeBackspace
    
  ' Überschriftenebene 1 löschen und danach durch Standard Wiki ersetzen
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Überschrift 1")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Überschrift 1")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Standard Wiki")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    
  ' Dokument speichern
    ChangeFileOpenDirectory "D:\tmp\"
    ActiveDocument.SaveAs2 FileName:=sHeadingNumber + " - " + sDocumentName + ".docx", FileFormat:= _
        wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, CompatibilityMode:=15
    ActiveDocument.Close
    Documents("Galeria-Anatomica-Wiki-Export.docx").Activate
    Selection.ExtendMode = False
    Selection.MoveDown Unit:=wdParagraph, Count:=1
    
  Next i
End Sub
Sub ListBookmarks()
'
' ListBookmarks Makro
'
'
   Dim i As Integer
   
   For i = 1 To ActiveDocument.Bookmarks.Count
     Debug.Print ActiveDocument.Bookmarks(i)
   Next i
   
End Sub

Sub GetHeadingNextText()

  Application.ScreenUpdating = False
  Dim RngHd As Range, h As Long, i As Long, strOut As String, ArrExpr()
  ArrExpr = Array("Abstract:", "Author:", "Keywords:", "References:", "Title:")

  For i = 0 To UBound(ArrExpr)
    strOut = strOut & vbCr & ArrExpr(i)
    For h = 3 To 4
      With ActiveDocument.Range
        With .Find
          .ClearFormatting
          .Replacement.ClearFormatting
          .Text = ArrExpr(i)
          .Style = "Überschrift " & h
          .Replacement.Text = ""
          .Forward = True
          .Wrap = wdFindStop
          .Format = True
          .MatchCase = True
          .MatchWildcards = False
          .MatchSoundsLike = False
          .MatchAllWordForms = False
          .Execute
        End With
        Do While .Find.Found
          Set RngHd = .Paragraphs.Last.Range.Next.Paragraphs.Last.Range
          With RngHd
            .End = .End - 1
            strOut = strOut & vbCr & .Text
          End With
          .Start = RngHd.End + 1
          .Find.Execute
        Loop
      End With
    Next
  Next
  
  Set RngHd = Nothing
  MsgBox "The following text is associated with -" & strOut
  Application.ScreenUpdating = True

End Sub
Sub CreateTableOfContent()
Attribute CreateTableOfContent.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.CreateTableOfContent"
'
' CreateTableOfContent Makro
'
'
    Selection.HomeKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdParagraph, Count:=1
    Selection.Style = ActiveDocument.Styles("Standard")
    With ActiveDocument
        .TablesOfContents.Add Range:=Selection.Range, RightAlignPageNumbers:= _
            False, UseHeadingStyles:=True, UpperHeadingLevel:=3, _
            LowerHeadingLevel:=3, IncludePageNumbers:=False, AddedStyles:="", _
            UseHyperlinks:=False, HidePageNumbersInWeb:=True, UseOutlineLevels:= _
            False
        .TablesOfContents(1).TabLeader = wdTabLeaderSpaces
        .TablesOfContents.Format = wdIndexIndent
    End With
    
End Sub
Sub DeleteYadaWikiLinks()
Attribute DeleteYadaWikiLinks.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.DeleteYadaWikiLinks"
'
' DeleteYadaWikiLinks Makro
'
'
  Dim i As Integer
    
  For i = 1 To 4410

    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("YadaWikiLink")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.Extend
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("YadaWikiLink")
    With Selection.Find
        .Text = "show="""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("YadaWikiLink")
    With Selection.Find
        .Text = """]"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.Delete Unit:=wdCharacter, Count:=1
    
  Next i
  
End Sub

Sub CreateBookmark()
'
' CreateBookmark Makro
'
'
    Dim sBookmark As String
    
    Selection.HomeKey Unit:=wdLine
    Selection.Extend
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    
    sBookmark = Selection.Text
    sBookmark = Replace(sBookmark, "_", "")
    sBookmark = Replace(sBookmark, " ", "_")
    sBookmark = Replace(sBookmark, "-", "_")
    sBookmark = Replace(sBookmark, ".", "_")
    sBookmark = Replace(sBookmark, "(", "_")
    sBookmark = Replace(sBookmark, ")", "_")
    sBookmark = Replace(sBookmark, ",", "_")
    sBookmark = Replace(sBookmark, "?", "_")
    sBookmark = Replace(sBookmark, "/", "_")
    sBookmark = Replace(sBookmark, "!", "_")
    
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:=sBookmark
        .DefaultSorting = wdSortByName
        .ShowHidden = False
    End With
    
    Selection.HomeKey Unit:=wdLine
    
End Sub
Sub CreateAudioLink()
Attribute CreateAudioLink.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.CreateAudioLink"
'
' CreateAudioLink Makro
'
'
  Dim i As Integer
  Dim sAudioFile As String
  Dim sLink As String
  
  For i = 1 To 104
  
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Kapitelinfo")
    With Selection.Find
        .Text = "Sound^t"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.ExtendMode = True
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Kapitelinfo")
    With Selection.Find
        .Text = "^l"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    If Len(Selection.Text) > 1 Then
      sAudioFile = Selection.Text
      Selection.ExtendMode = False
      Selection.HomeKey Unit:=wdLine
      Selection.MoveDown Unit:=wdParagraph, Count:=1
      sLink = "<audio id=" + Chr$(34) + "audio-player" + Chr$(34) + " controls src=" + Chr$(34) + "http://192.168.178.65:9980/wp-content/uploads/" + Left$(sAudioFile, 7) + "-1.wav" + Chr$(34) + " type=" + Chr$(34) + "audio/wav" + Chr$(34) + "></audio>"
      Selection.Text = sLink + vbCrLf
    End If
  
 Next i
  
End Sub

Sub ListAllBuiltInStyleNames()
    'Create a new, blank document and insert names of all built-in styles
    Dim oSty As Style
    Dim oDoc As Document
    Dim n As Long
    
    n = 0
    Set oDoc = Documents.Add
    With oDoc
        .Range.Text = ""
        For Each oSty In .Styles
            If oSty.BuiltIn = True Then
                n = n + 1
                .Range.InsertAfter oSty.NameLocal & vbCr
            End If
        Next oSty
    End With
    
    MsgBox n & " built-in styles found.", vbOKOnly, "Names of All Built-in Style Names"
    
    Set oDoc = Nothing

End Sub
Sub DeleteHiddenStyles()
Attribute DeleteHiddenStyles.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.DeleteHiddenStyles"
'
' DeleteHiddenStyles Makro
'
'
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Kapitelinfo")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Keywords")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Verweis")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub PrepareExport()
'
' PrepareExport Makro
'
'
    DeleteHiddenStyles
    DeleteAllBookmarks
    DeleteUnusedStyles
    
    MsgBox "Fertig", vbOKOnly, "Prepare Export"

End Sub
Sub GetHeadingNumber()
'
' GetHeadingNumber Makro

  MsgBox Selection.Range.ListFormat.ListString


End Sub
Sub Makro1()
Attribute Makro1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Makro1"
'
' Makro1 Makro
'
'
    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = wdNormalView
    Else
        ActiveWindow.View.Type = wdNormalView
    End If
    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    Else
        ActiveWindow.View.Type = wdPrintView
    End If
End Sub
Sub InsertImage()
Attribute InsertImage.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.InsertImage"
'
' InsertImage Makro
'
'
    Dim sImageName As String
    Dim sFieldText As String
    Dim bFound As Boolean
    
    bFound = True
    
    Selection.Find.ClearFormatting
    
    Do While bFound
        Selection.Find.Style = ActiveDocument.Styles("Bild")
        With Selection.Find
            .Text = "/uploads/"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        bFound = Selection.Find.Execute
        If bFound Then
            Selection.MoveRight Unit:=wdCharacter, Count:=1
            Selection.ExtendMode = True
            Selection.Find.ClearFormatting
            Selection.Find.Style = ActiveDocument.Styles("Bild")
            With Selection.Find
                .Text = "#"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindStop
                .Format = True
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.Execute
            Selection.ExtendMode = False
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            sImageName = Selection.Text
            If LCase(Right(sImageName, 3)) = "jpg" Or LCase(Right(sImageName, 3)) = "gif" Then
                Selection.MoveUp Unit:=wdParagraph, Count:=1
                Selection.TypeParagraph
                Selection.MoveUp Unit:=wdParagraph, Count:=1
                sImageName = "D:\\Projekte\\Galeria Anatomica\\Wiki\\Grafiken\\Diverse\\" + sImageName
                sFieldText = "INCLUDEPICTURE " + Chr$(34) + sImageName + Chr$(34) + " \d"
                Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:=sFieldText, PreserveFormatting:=False
                Selection.Style = ActiveDocument.Styles("GrafikEingebunden")
                Selection.MoveDown Unit:=wdParagraph, Count:=2
             Else
                Selection.MoveDown Unit:=wdParagraph, Count:=1
             End If
         End If
     Loop
     
End Sub

Sub LinkImagesToProducts()
'
' LinkImagesToProducts Makro
'
'
  Const MAX_PICTURES = 92
  
  Dim i, iPos As Integer
  Dim bFound As Boolean
  Dim sLine As String
  Dim sLineItems() As String
  Dim jpg_name(1 To MAX_PICTURES) As String
  Dim article_id(1 To MAX_PICTURES) As String
  Dim product_name(1 To MAX_PICTURES) As String
  Dim description(1 To MAX_PICTURES) As String
  Dim sHTML As String
  Dim qm As String
  
  qm = Chr$(34)

  Open "d:\projekte\galeria-anatomica\bilder\wiki-image-to-wc-product-mapping.csv" For Input As #1
  
  i = 1
  While Not EOF(1)
    Line Input #1, sLine
    sLineItems = Split(sLine, Chr$(9))
    jpg_name(i) = sLineItems(0)
    article_id(i) = Left$(sLineItems(1), Len(sLineItems(1)) - 4)
    product_name(i) = sLineItems(2)
    description(i) = sLineItems(3)
    i = i + 1
  Wend
  
  Close #1
    
  Selection.Find.ClearFormatting
  bFound = True
  
  Do While bFound
      
      Selection.Find.Style = ActiveDocument.Styles("Bild")
      
      With Selection.Find
          .Text = ""
          .Replacement.Text = ""
          .Forward = True
          .Wrap = wdFindStop
          .Format = True
          .MatchCase = False
          .MatchWholeWord = False
          .MatchWildcards = False
          .MatchSoundsLike = False
          .MatchAllWordForms = False
      End With
      
      bFound = Selection.Find.Execute
      
      If bFound Then
          For i = 1 To MAX_PICTURES
              If InStr(Selection.Text, jpg_name(i)) Then
                  sHTML = "<a href=" + qm + "http://192.168.178.61:9980/produkt/" + product_name(i) + "/" + qm + ">" + _
                            "<img class=" + qm + "product" + qm + " " + _
                              "title=" + qm + description(i) + qm + " " + _
                              "src=" + qm + "http://192.168.178.61:9980/wp-content/uploads/wiki/" + jpg_name(i) + qm + " " + _
                              "alt=" + qm + description(i) + qm + _
                            "/>" + _
                          "</a>"
                  sHTML = Replace(sHTML, "<", "+++")
                  sHTML = Replace(sHTML, ">", "---")
                  sHTML = Replace(sHTML, qm, "#")
                  Selection.Text = sHTML + vbNewLine
                  Selection.Style = ActiveDocument.Styles("ProduktBild")
                  Selection.MoveDown Unit:=wdParagraph, Count:=1
                  Exit For
              End If
          Next i
      End If
   
   Loop

End Sub


