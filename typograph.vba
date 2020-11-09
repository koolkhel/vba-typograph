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

Sub Typograph()
    '
    ' Typograph Макрос
    '
    '
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
End If
'восстановить пользовательскую установку
Options.AutoFormatAsYouTypeReplaceQuotes = blnQuotes
End Sub

