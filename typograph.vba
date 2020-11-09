Sub ReplaceQuotes(toReplace As String)

    With Selection.Find
       .ClearFormatting
       .Replacement.ClearFormatting
       .Text = toReplace
       With .Replacement
       .Text = "«\1»"
       End With
       .Wrap = wdFindStop
       .MatchWildcards = True
       .Replacement.Font.Italic = True
       .Execute Replace:=wdReplaceAll
    End With

End Sub

Sub FixInnerQuotes()

    Dim txt As String
    Dim n As Long
    Dim insideQuotes As Integer
    Dim c As String
    Dim result As String
    
    insideQuotes = 0
    txt = Selection.Range.Text
    
    For n = 1 To Len(txt)
    
        c = Mid(txt, n, 1)
        
        If c = "«" Then
            insideQuotes = insideQuotes + 1
            
            If insideQuotes = 2 Then
                result = result & "„"
            ElseIf insideQuotes = 1 Then
                result = result & "«"
            End If
            
        ElseIf c = "»" Then
            insideQuotes = insideQuotes - 1
            
            If insideQuotes = 1 Then
                result = result & "“"
            ElseIf insideQuotes = 0 Then
                result = result & "»"
            End If
            
        Else
            result = result & c
        End If
    
    Next n
    
    Selection.Range.FormattedText.Text = result

End Sub

Sub Typograph()

Dim blnQuotes As Boolean
'запомнить пользовательскую установку
blnQuotes = Options.AutoFormatAsYouTypeReplaceQuotes
Options.AutoFormatAsYouTypeReplaceQuotes = False
' прямые кавычки
If Selection.Type = wdSelectionNormal Then
    ReplaceQuotes ("""(*)""")
    ReplaceQuotes ("“(*)”")
    ReplaceQuotes ("„(*)“")
    ReplaceQuotes ("«(*)»")
    ' тире (хоть тут по тексту и не видно
    With Selection.Find
       .ClearFormatting
       .Replacement.ClearFormatting
       .Text = " - "
       .Replacement.Text = " – "
       .Wrap = wdFindContinue
       .MatchWildcards = True
       .Execute Replace:=wdReplaceAll
    End With
    
    FixInnerQuotes
    
End If
'восстановить пользовательскую установку
Options.AutoFormatAsYouTypeReplaceQuotes = blnQuotes
End Sub

