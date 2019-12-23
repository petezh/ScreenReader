
    Dim p
    Dim pCount As Long

    pCount = 0

    'Make sure status bar is visible for progress indicator
    Application.StatusBar = True

    'Loop each paragraph
    For Each p In ActiveDocument.Paragraphs
        pCount = pCount + 1
        Application.StatusBar = "Processing paragraph " & pCount & " of " & ActiveDocument.Paragraphs.Count

        'Select each non-blank body text paragraph
        If p.OutlineLevel = wdOutlineLevelBodyText And Len(p) > 1 Then
            p.Range.Select

            'Highlight the cites so they don't disappear
            With Selection.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = ""
                .Wrap = wdFindStop
                .Replacement.Text = ""
                .Format = True
                .Style = "Cite"
                .Execute

                'Skip the paragraph if cite is found
                If .Found = True Then GoTo Skip
            End With

            'Select the paragraph, shorten to keep line breaks
            p.Range.Select
            Selection.MoveEndWhile Cset:=vbCrLf, Count:=-1
            Selection.MoveEndWhile Cset:=" ", Count:=-1
            Selection.MoveStartWhile Cset:=vbCrLf, Count:=1
            Selection.MoveStartWhile Cset:=" ", Count:=1

            'Delete all non-highlighted text
            With Selection.Find
                .ClearFormatting
                .Wrap = wdFindStop
                .MatchWildcards = True
                .Format = True
                .Highlight = False
                .ParagraphFormat.OutlineLevel = wdOutlineLevelBodyText
                With .Replacement
                    .ClearFormatting
                    .Style = "Underline"
                    .Highlight = True
                    .Text = " "
                End With
                .Execute Replace:=wdReplaceAll
            End With



        End If
Skip:
    Next p

    'Clean up and supress errors
    Selection.Find.ClearFormatting
    Selection.Find.MatchWildcards = False
    Selection.Find.Replacement.ClearFormatting

    ActiveDocument.ShowGrammaticalErrors = False
    ActiveDocument.ShowSpellingErrors = False


