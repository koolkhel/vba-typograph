Option Explicit

Sub FormatText()
    Dim dict As New Dictionary
    
    Dim pClient As WebClient
    Dim pRequest As New WebRequest
    Dim pResponse As WebResponse
    
    Dim str_boundary As String
    Dim str_text As String
    Dim str_encoding As String
    
    Set pClient = New WebClient
    
    pClient.BaseUrl = "http://typograph.bitterman.ru:8000"
    
    str_boundary = RandomString(24)
    str_text = Selection.Text
    #If Mac Then
    str_encoding = "maccyrillic"
    #Else
    str_encoding = "cp1251"
    #End If
    
    pRequest.Resource = "/"
    pRequest.Method = HttpPost
    pRequest.ResponseFormat = PlainText
        
    pRequest.ContentType = lbb_make_post_cont_type(str_boundary)
    pRequest.Body = lbb_make_post_body(str_text, str_encoding, str_boundary)
     
     ' VBA-Web seems to say more to Content-Length than there really is
     pRequest.ContentLength = CStr(CInt(pRequest.ContentLength) - 6)
     ' WebHelpers.EnableLogging = True
     
    Set pResponse = pClient.Execute(pRequest)
    If pResponse.StatusCode = WebStatusCode.Ok Then
        Selection.Text = pResponse.Content & vbNewLine
    End If
End Sub

Public Function lbb_make_post_body(str_xml As String, encoding As String, str_boundary As String)
    
    lbb_make_post_body = "" _
    & "----------------------------" & str_boundary & vbNewLine _
    & "Content-Disposition: form-data; name='text'; filename='test.txt'" & vbNewLine _
    & "Content-Type: text/plain" & vbNewLine & vbNewLine _
    & str_xml & vbNewLine _
    & "----------------------------" & str_boundary & vbNewLine _
    & "Content-Disposition: form-data; name='encoding'" & vbNewLine & vbNewLine _
    & encoding & vbNewLine _
    & "----------------------------" & str_boundary & "--" & vbNewLine

' "payload" is the name of the parameter, the server is watching for files,
' don't know, what "filename" is used for

End Function

Public Function lbb_make_post_cont_type(str_boundary As String)

    lbb_make_post_cont_type = "multipart/form-data; boundary=--------------------------" & str_boundary

End Function

' ## taken from https://www.thespreadsheetguru.com/the-code-vault/generate-string-random-characters-vba-codet
' but removed the special characters from the Char Array

Function RandomString(Length As Integer)
'PURPOSE: Create a Randomized String of Characters
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim CharacterBank As Variant
Dim x As Long
Dim str As String

'Test Length Input
  If Length < 1 Then
    MsgBox "Length variable must be greater than 0"
    Exit Function
  End If

CharacterBank = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", _
  "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", _
  "y", "z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", _
  "A", "B", "C", "D", "E", "F", "G", "H", _
  "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", _
  "W", "X", "Y", "Z")
  

'Randomly Select Characters One-by-One
  For x = 1 To Length
    Randomize
    str = str & CharacterBank(Int((UBound(CharacterBank) - LBound(CharacterBank) + 1) * Rnd + LBound(CharacterBank)))
  Next x

'Output Randomly Generated String
  RandomString = str
  
End Function

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
    Dim blnQuotes As Boolean
    '  запомнить пользовательскую установку
    blnQuotes = Options.AutoFormatAsYouTypeReplaceQuotes
    Options.AutoFormatAsYouTypeReplaceQuotes = False
    ' прямые кавычки
    If Selection.Type = wdSelectionNormal Then
        ' Selection.ClearFormatting
        FormatText
        ReplaceQuotes ("«(*)»")
    End If
'восстановить пользовательскую установку
Options.AutoFormatAsYouTypeReplaceQuotes = blnQuotes
End Sub

Sub PasteNoFormat()
    Selection.PasteAndFormat wdFormatPlainText
End Sub
