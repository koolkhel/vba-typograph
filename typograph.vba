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
    With Selection.Find
       .ClearFormatting
       .Replacement.ClearFormatting
       .Text = """(*)"""
       .Replacement.Text = "«\1»"
       .Wrap = wdFindContinue
       .MatchWildcards = True
       .Replacement.Font.Italic = True
       .Execute Replace:=wdReplaceOne
    End With
    ' лапки
    With Selection.Find
       .ClearFormatting
       .Replacement.ClearFormatting
       .Text = "“(*)”"
       .Replacement.Text = "«\1»"
       .Wrap = wdFindContinue
       .MatchWildcards = True
       .Replacement.Font.Italic = True
       .Execute Replace:=wdReplaceOne
    End With
    ' лапки 2
    With Selection.Find
       .ClearFormatting
       .Replacement.ClearFormatting
       .Text = "„(*)“"
       .Replacement.Text = "«\1»"
       .Wrap = wdFindContinue
       .MatchWildcards = True
       .Replacement.Font.Italic = True
       .Execute Replace:=wdReplaceOne
    End With
    ' елочки - курсив
    With Selection.Find
       .ClearFormatting
       .Replacement.ClearFormatting
       .Text = "«(*)»"
       .Replacement.Text = "«\1»"
       .Wrap = wdFindContinue
       .MatchWildcards = True
       .Replacement.Font.Italic = True
       .Execute Replace:=wdReplaceOne
    End With
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

