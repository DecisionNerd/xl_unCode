Attribute VB_Name = "Module1"
Sub xl_unCode()

'use of thw Wolfram|Alpha API requires compliance with their Terms of Use
'

'MUST ENABLE references: Tools > References... (Microsoft HTML Object Library & Microsoft Internet Controls)

'select country_name from page
Dim Last As Integer
Last = ActiveSheet.Range("A10000").End(xlUp).Row
If Last = 1 Then
  MsgBox "no records to process in column A"
  Exit Sub
End If

MsgBox "number of records to process: " & Last - 1

'Wolfram Alpha appid
Dim appid As String
appid = "xxxxxx-xxxxxxxxxx"

Dim cName As String
Dim i As Integer
For i = 2 To Last
  cName = ActiveSheet.Range("a" & i).Value
  'MsgBox "2.cName: " & cName
  cCode = ActiveSheet.Range("b" & i).Value
    'get UN code from Wolfram Alpha
    Dim url As String
    url = "http://api.wolframalpha.com/v2/query?appid=" & appid & "&input=un%20code%20" & cName & "&format=plaintext"
    'MsgBox "3.URL: " & url
    Dim IE As New InternetExplorer
    IE.Visible = False
        IE.navigate url
    Do
      DoEvents
    Loop Until IE.readyState = READYSTATE_COMPLETE
    
    'find answer in page
    Dim Doc As HTMLDocument
    Set Doc = IE.document
    Dim tags As String
    tags = Trim(Doc.getElementsByTagName("plaintext")(1).innerText)
    'MsgBox "5.plaintext: " & tags
    Dim noTAGS As String
    noTAGS = Mid(tags, 12, 3)
    'output to cell
    'MsgBox "6.UN CODE: " & noTAGS
    ActiveSheet.Range("b" & i).Value = noTAGS
    
Next i

MsgBox "unCode DONE using the Wolfram|Alpha API.  This macro is experimental so review results - particularly for Korea, Congo, and former countries such as the USSR and Yugoslavia"
    
End Sub
