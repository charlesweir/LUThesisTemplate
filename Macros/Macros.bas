Attribute VB_Name = "NewMacros"
' General macros for use by Charles

Sub UpdateAllFieldsIn(doc As Document)
' Updates all fields, and tables of contents.
' Also updates Mendeley references and sets the table of references style (for IEEE) to MendeleyReference
    Application.StatusBar = "Updating fields..." 'N.b. Doesn't seem to work.

    ' Do this twice. Figure numbers seem to update the first time, references to them the second time
    Dim i As Long
    For i = 1 To 2
        '' Update tables. We do this first so that they contain all necessary
        '' entries and so extend to their final number of pages.
        Dim toc As TableOfContents
        For Each toc In doc.TablesOfContents
            toc.Update
        Next toc
        Dim tof As TableOfFigures
        For Each tof In doc.TablesOfFigures
            tof.Update
        Next tof
        '' Update fields everywhere. This includes updates of page numbers in
        '' tables (but would not add or remove entries). This also takes care of
        '' all index updates.
        Dim sr As Range
        For Each sr In doc.StoryRanges
            sr.Fields.Update
            While Not (sr.NextStoryRange Is Nothing)
                Set sr = sr.NextStoryRange
                '' FIXME: for footnotes, endnotes and comments, I get a pop-up
                '' "Word cannot undo this action. Do you want to continue?"
                sr.Fields.Update
            Wend
        Next sr
        
    Next i
    
    ' Mendeley refresh. Needs Tools - References - MendeleyPlugin ticked.
    Refresh
    
    ReformatMendeleyReferenceList
End Sub
'' Update all the fields, indexes, etc. in the active document.
'' This is a parameterless subroutine so that it can be used interactively.
Sub UpdateAllFields()
    UpdateAllFieldsIn ActiveDocument
End Sub



Sub ZoomToPreferred(Dummy)
Attribute ZoomToPreferred.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Zoom120"
'
' Set document to preferred zoom level.
'
'
    ActiveWindow.ActivePane.View.Zoom.Percentage = 200
End Sub
Sub SetupMasterForEditing()
Attribute SetupMasterForEditing.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.SetupMasterForEditing"
'
' The document comes up in a default mode with subdocuments locked and showing comments. Change it back.
'
    ChangeView (wdOutlineView)
    ActiveDocument.Subdocuments.Expanded = True
    ChangeView (wdPrintView)
    With ActiveWindow.View
        .ShowRevisionsAndComments = False
        .RevisionsView = wdRevisionsViewFinal
    End With
    ActiveWindow.DocumentMap = True 'And restore the Navigation Pane
    ZoomToPreferred (0)
End Sub
Sub ChangeView(View As Integer)
' Change to view View - see https://msdn.microsoft.com/en-us/library/office/ff836365.aspx for values
    ActiveWindow.ActivePane.View.Type = View
    DoEvents
End Sub


Sub ChangeCaseAllTitles()
    ChangeCaseForParticularTitleStyle ("Heading 1")
    ChangeCaseForParticularTitleStyle ("Heading 2")
    ChangeCaseForParticularTitleStyle ("Heading 3")
End Sub

Sub ChangeCaseCurrentSelection()
    ChicagoTitleCase ("")
End Sub

Sub ChangeCaseForParticularTitleStyle(style As String)
    With Selection.Find
         .ClearFormatting
         .Wrap = wdFindContinue
         .Forward = True
         .Format = True
         .MatchWildcards = False
         .Text = ""
         .style = ActiveDocument.Styles(style)
         .Execute
         While .Found
             ChicagoTitleCase ("")
             Selection.Collapse Direction:=wdCollapseEnd
            .Execute
         Wend
     End With
End Sub

Sub ChicagoTitleCase(Dummy As String)
' Headline-Style Capitalization (according to the Chicago Manual of Style)
' ? Capitalize: ?first and last word, first word after a colon (subtitle)
' all major Words(nouns, pronouns, verbs, adjectives, adverbs)

' lowercase: articles(the, a, an)
'   ?prepositions (regardless of length)
'   ?conjunctions (and, but, for, or, nor)
'   ?to, as

' Hyphenated; Compounds: Print always; capitalize; First; element
' ?lowercase second element for articles, prepositions, conjunctions and if the first element is a prefix or combining form that could not stand by itself (unless the second element is a proper noun / proper adjective)


    Dim lclist As String
    Dim wrd As Integer
    Dim sTest As String

    ' list of lowercase words, surrounded by spaces
    lclist = " the a an and but for or nor as " + "about above across after against around at before behind below beneath beside besides between beyond by down during except for from in inside into like near of off on out outside over since through throughout till to toward under until up upon versus with without "
    
    Selection.Range.Case = wdTitleWord

    For wrd = 2 To (Selection.Range.Words.Count - 2)
        sTest = Trim(Selection.Range.Words(wrd))
        sTest = " " & LCase(sTest) & " "
        If InStr(lclist, sTest) Then
            Selection.Range.Words(wrd).Case = wdLowerCase
        End If
    Next wrd
End Sub

Sub SavePDF()
Attribute SavePDF.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
' Save the current document as a PDF in the same directory with the PDF document extension. Not on Mac 2016

    ' Ensure we're saving fonts as well.
    ActiveDocument.SaveSubsetFonts = True
    ActiveDocument.EmbedTrueTypeFonts = True
    
    Dim baseName As String
    baseName = ActiveDocument.Path & "\" & Left(ActiveDocument.name, InStrRev(ActiveDocument.name, ".") - 1)
    
    ActiveDocument.ExportAsFixedFormat OutputFileName:= _
        baseName + ".pdf", _
        ExportFormat:=wdExportFormatPDF, OpenAfterExport:=True, OptimizeFor:= _
        wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=1, To:=1, _
        Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateHeadingBookmarks, DocStructureTags:=True, _
        BitmapMissingFonts:=True, UseISO19005_1:=False
        
    ' And reset back to normal afterwards:
    ActiveDocument.SaveSubsetFonts = False
    ActiveDocument.EmbedTrueTypeFonts = False
    
End Sub

Sub PasteQuote()
Attribute PasteQuote.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
'
' Paste a quote from NVivo, reformatting.
' I can't figure out how Range objects work.
' No matter how much I paste into the range, (and edit it) it still seems 0 characters long and only to contain the
' position before all pasted text.
'
    Dim quoteRange As Range
    Selection.TypeText (" (P)")
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    ActiveDocument.Bookmarks.Add name:="temp", Range:=Selection.Range
    Selection.MoveLeft Unit:=wdCharacter, Count:=3
    
    Set quoteRange = Selection.Range
    With quoteRange
        
        .PasteAndFormat (wdFormatPlainText)
        .style = ActiveDocument.Styles("Quote")
        'Remove double spaces after full stops:
        
        'With quoteRange.Find
         '   .ClearFormatting
          '  .Replacement.ClearFormatting
           ' .Text = ".  "
            '.Replacement.Text = ". "
        '    .Forward = True
         '   .Wrap = wdFindStop
          '  .Format = False
           ' .MatchCase = False
       '     .MatchWholeWord = False
        '    .MatchWildcards = False
         '   .MatchSoundsLike = False
          '  .MatchAllWordForms = False
       '     .Execute Replace:=wdReplaceAll
            
            '.Text = "^p"
            '.Replacement.Text = " "
            '.Execute Replace:=wdReplaceAll
        '    End With 'Find
        
        .AutoFormat
    End With 'quoteRange
    ActiveDocument.Bookmarks("temp").Select
    ActiveDocument.Bookmarks(Index:="temp").Delete
    
End Sub
Sub UncurlQuoteMarks()
Attribute UncurlQuoteMarks.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.QuoteSort"
'
' Uncurls quote marks in the current selection
'
    Dim PrevOption As Boolean
    
    With Options
        PrevOption = .AutoFormatAsYouTypeReplaceQuotes
        .AutoFormatAsYouTypeReplaceQuotes = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "'"
        .Replacement.Text = "'"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = """"
        .Replacement.Text = """"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    With Options
        .AutoFormatAsYouTypeReplaceQuotes = PrevOption
    End With
End Sub

Sub ReformatMendeleyReferenceList()
'
' MendeleyRestyle Macro
' Selects IEEE formated references and changes their style to MendeleyReference.
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.style = ActiveDocument.Styles("MendeleyReference")
    With Selection.Find
        .Text = "\[[0-9]@\]^t*^13"
        .Replacement.Text = "^&"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
End Sub

Sub AFormatMyPicture()
' Resets the layout of the selected image or text frame:
'   In two column mode, to be top or bottom of its column or page.
'   In single column mode, float top or bottom of the page if large
'    Else Float Left, Right, Above and Below the anchor

    Dim myShape As Shape
      Dim AnchorParagraph As Paragraph
      
    If Selection.ShapeRange.Count > 0 Then
        Set myShape = Selection.ShapeRange(1)
    Else
        MsgBox "Please select a float picture first."
        Exit Sub
    End If

    With myShape
        If Selection.Sections(1).PageSetup.TextColumns.Count > 1 Then
            ' Two columns. In column if small enough, else page. Toggle top/bottom
            MaxSingleColumnImageWidth = Selection.Sections(1).PageSetup.TextColumns.Width + Selection.Sections(1).PageSetup.TextColumns.Spacing
            .WrapFormat.Type = wdWrapTopBottom
            If .Width > MaxSingleColumnImageWidth Then
                .RelativeHorizontalPosition = wdRelativeHorizontalPositionMargin
            Else
                .RelativeHorizontalPosition = wdRelativeHorizontalPositionColumn
            End If
            .RelativeVerticalPosition = wdRelativeVerticalPositionMargin
            .Left = wdShapeCenter
            If .Top = wdShapeTop Then
                .Top = wdShapeBottom
            Else
                .Top = wdShapeTop
            End If
        Else
            ' One column.
            HalfPageWidth = Selection.Sections(1).PageSetup.TextColumns.Width / 2
            If .Width < HalfPageWidth Then
                'Small picture. Put below anchor, wrap around. Toggle left/right, above below
                .WrapFormat.Type = wdWrapSquare
                .RelativeHorizontalPosition = wdRelativeHorizontalPositionColumn
                .RelativeVerticalPosition = wdRelativeVerticalPositionParagraph
                If .Left = wdShapeRight And .Top = 0 Then
                    pos = 2 ' Below Left
                    .Left = wdShapeLeft
                ElseIf .Left = wdShapeLeft And .Top = 0 Then
                    pos = 3 ' Above Right
                    .Left = wdShapeRight
                ElseIf .Left = wdShapeRight And .Top < 0 Then
                    pos = 4 'Above Left
                    .Left = wdShapeLeft
                Else
                    pos = 1 'Below right
                    .Left = wdShapeRight
                    .Top = 0
                End If
                
                If pos >= 3 Then
                    ' Locate the image just above the anchor
                    Set AnchorParagraph = .Anchor.Paragraphs(1)
                    ParaSpacing = AnchorParagraph.SpaceAfter
                    .RelativeVerticalPosition = wdRelativeVerticalPositionPage
                    .Top = AnchorParagraph.Range.Information(wdVerticalPositionRelativeToPage) - .Height - ParaSpacing
                    Do Until .Top + .Height + 1 + ParaSpacing > AnchorParagraph.Range.Information(wdVerticalPositionRelativeToPage)
                        .IncrementTop (1)
                    Loop
                End If
            Else
                ' Big picture: Toggle top/bottom of page
                .WrapFormat.Type = wdWrapTopBottom
                .RelativeVerticalPosition = wdRelativeVerticalPositionMargin
                .RelativeHorizontalPosition = wdRelativeHorizontalPositionColumn
                .Left = wdShapeCenter
                If .Top = wdShapeTop Then
                    .Top = wdShapeBottom
                Else
                    .Top = wdShapeTop
                End If
            End If
        End If
            
    End With

End Sub

Sub ImageDownABit()
    ' Take the currently selected image down a pixel without altering formatting
    Dim myShape As Shape
      
    If Selection.ShapeRange.Count > 0 Then
        Set myShape = Selection.ShapeRange(1)
    Else
        MsgBox "Please select a float picture first."
        Exit Sub
    End If

    myShape.IncrementTop (1)
End Sub
