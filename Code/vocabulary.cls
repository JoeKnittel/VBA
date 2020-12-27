VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'requires that all vars be explicitly declared
Option Explicit

'performs api request at the desired url, and outputs json data
Function getData(myUrl As String) As String
    
    Dim winHttpReq As Object
    Set winHttpReq = CreateObject("Microsoft.XMLHTTP")
    
    winHttpReq.Open "GET", myUrl, False
    winHttpReq.Send
    
    getData = winHttpReq.ResponseText

End Function

'gets the url of a searched word
Function getURL(word As String) As String
    
    Dim p1, p2, p3, apiKey, url As String
    
    p1 = "https://www.dictionaryapi.com/api/v3/references/collegiate/json/"
    p2 = word
    p3 = "?key="
    
    apiKey = "7a554f7f-1b41-4f71-be12-d996e7ffd9fd"
    
    url = p1 & p2 & p3 & apiKey
    
    getURL = url

End Function

'determines if the inputted word is in the dictionary
Function isWord(word As String) As Boolean
    
    Dim Json As Object
    Set Json = JsonConverter.ParseJson(getData(getURL(word)))
    
    If Len(JsonConverter.ConvertToJson(Json)) > 500 Then
        isWord = True
    Else
        isWord = False
    End If
        
End Function

'debug function: tests if a specified word is found in the dictionary
Sub isWordCheck()
    
    Debug.Print (isWord("test"))

End Sub

'automates actions (e.g., prints part of speech and definition) when a word is added/subtracted from list
Private Sub Worksheet_Change(ByVal Target As Range)
    
    Dim temp As String
    
    If Target.Column = "2" And Target.Value <> "" Then
        
        If isWord(Target.Value) Then
            
            Call printPoS(Target.row, 3)
            Call printDef(Target.row, 4)
            Call printPronunciation(Target.row, 5)
        
            If Cells(Target.row, 2) <> getID(Cells(Target.row, 2).Value) Then
                Call transformInput(Target.row, 2)
            End If
        Else
            Target.Value = ""
        End If
        
    ElseIf Target.Column = "2" And Target.Value = "" Then
        
        Cells(Target.row, Target.Column + 1).Value = ""
        Cells(Target.row, Target.Column + 2).Value = ""
        Cells(Target.row, Target.Column + 3).Value = ""
    
    End If

End Sub

'munges json data from api response and finds the desired word's definition
Function getDef(word As String) As String
    
    Dim data As String
    data = getData(getURL(word))
    Dim Json As Object
    Set Json = JsonConverter.ParseJson(data) 'this is the data, parsed
    
    If Left(JsonConverter.ConvertToJson(Json(1)("def")(1)("sseq")(1)(1)(2), Whitespace:=2), 1) <> "[" Then
    
        getDef = JsonConverter.ConvertToJson(Json(1)("def")(1)("sseq")(1)(1)(2)("dt")(1)(2))
    
    Else
    
        getDef = JsonConverter.ConvertToJson(Json(1)("shortdef")(1))
    
    End If

End Function

'munges json data from the api response and finds the desired word's id
Function getID(word As String) As String
    
    Dim data, temp As String
    Dim colonPos As Integer
    data = getData(getURL(word))
    Dim Json As Object
    Set Json = JsonConverter.ParseJson(data)
    
    temp = JsonConverter.ConvertToJson(Json(1)("meta")("id"))
    temp = Mid(temp, 2, Len(temp) - 2)
    
    colonPos = InStr(temp, ":")
    
    If colonPos > 1 Then
        
        temp = Left(temp, colonPos - 1)
    
    End If

    getID = temp
    
End Function


'munges json data from api response and finds the desired word's part of speech
Function getPoS(word As String) As String
    
    Dim data, rawPoS As String
    data = getData(getURL(word))
    Dim Json As Object
    Set Json = JsonConverter.ParseJson(data)
    
    rawPoS = JsonConverter.ConvertToJson(Json(1)("fl"))
    getPoS = Mid(rawPoS, 2, Len(rawPoS) - 2)

End Function

'munges json data from api response and finds desired the url of the desired word's pronunciation
Function getPronunciation(word As String) As String
    
    Dim data, temp, folder As String
    data = getData(getURL(word))
    Dim Json As Object
    Set Json = JsonConverter.ParseJson(data)
    
    temp = JsonConverter.ConvertToJson(Json(1)("hwi")("prs")(1)("sound")("audio"))
    temp = Mid(temp, 2, Len(temp) - 2) 'removes quotations
    
    If IsNumeric(Left(temp, 1)) Then
        folder = "number"
    Else
        folder = Left(temp, 1)
    End If
    
    temp = "https://media.merriam-webster.com/audio/prons/en/us/mp3/" & folder & "/" & temp & ".mp3" 'compose full url
    
    getPronunciation = temp
    
End Function

'debug function: prints the url of the word's pronunciation file
Sub outputPronunciation()
    
    Debug.Print (getPronunciation("tester"))

End Sub

'debug function: prints the definition of a pre-specified word, in its raw form
Sub outputRawDef()
    
    Debug.Print (getDef("tester"))

End Sub

'debug function: prints the definition of a transformed, pre-specified word
Sub outputDef()
    
    Debug.Print (getTransformedDef("tester"))

End Sub

'cleans up definition string (i.e., formatting tags), preparing it for printing to screen
Function getTransformedDef(word As String) As String
    
    'removes " from start and end of def (and {bc} from start, if necessary)
    Dim temp As String
    temp = getDef(word)
    If Mid(temp, 2, 3) = "{bc" Then
        temp = Mid(getDef(word), 6, Len(getDef(word)) - 6)
    Else
        temp = Mid(getDef(word), 2, Len(getDef(word)) - 2)
    End If
    
    'removes {bc} tag, if present
    Dim posBC As Integer
    posBC = InStr(temp, "{bc}")
    If posBC > 0 Then
        temp = Left(temp, posBC - 1) & ": " & Mid(temp, posBC + 4, (Len(temp) - posBC) - 3)
    End If
    
    'removes {sx} tag(s), if present
    Dim posSX, endSX, wordLen As Integer
    posSX = InStr(temp, "{sx|")
    endSX = InStr(temp, "||")
    wordLen = (endSX - posSX) - 4
    While posSX > 0
        temp = Left(temp, posSX - 1) & Mid(temp, posSX + 4, wordLen) & Right(temp, Len(temp) - posSX - 6 - wordLen)
        posSX = InStr(temp, "{sx|")
        endSX = InStr(temp, "||")
        wordLen = (endSX - posSX) - 4
    Wend
    
    'removes {a_link} tag(s), if present
    Dim posALink, endALink, linkLen As Integer
    posALink = InStr(temp, "{a_link")
    While posALink > 0
        endALink = InStr(Mid(temp, posALink), "}") + posALink - 1
        linkLen = (endALink - posALink) - 8
        temp = Left(temp, posALink - 1) & Mid(temp, posALink + 8, linkLen) & Mid(temp, endALink + 1)
        posALink = InStr(temp, "{a_link")
    Wend
    
    'removes {d_link} tag(s), if present
    Dim posDLink, endDLink, dlinkLen, tagEnd As Integer
    posDLink = InStr(temp, "{d_link")
    While posDLink > 0
        endDLink = InStr(Mid(temp, posDLink + 8), "|") + posDLink + 7
        dlinkLen = (endDLink - posDLink) - 8
        tagEnd = InStr(Mid(temp, posDLink), "}") + posDLink - 1
        temp = Left(temp, posDLink - 1) & Mid(temp, posDLink + 8, dlinkLen) & Mid(temp, tagEnd + 1)
        posDLink = InStr(temp, "{d_link")
    Wend
    
    'removes {it} tags, if present
    Dim posIt, endIt As Integer
    posIt = InStr(temp, "{it")
    While posIt > 0
        endIt = InStr(temp, "{/it")
        temp = Left(temp, posIt - 1) & Mid(temp, posIt + 4, (endIt - posIt) - 4) & Mid(temp, endIt + 5)
        posIt = InStr(temp, "{it")
    Wend
    
    'removes {b} tags, if present
    Dim posB, endB As Integer
    posB = InStr(temp, "{b")
    While posB > 0
        endB = InStr(temp, "{/b")
        temp = Left(temp, posB - 1) & Mid(temp, posB + 3, (endB - posB) - 3) & Mid(temp, endB + 4)
        posB = InStr(temp, "{b")
    Wend
    
    'removes {dx_def} tags, if present
    Dim posDX, endDX As Integer
    posDX = InStr(temp, "{dx_def")
    endDX = InStr(temp, "{/dx_def")
    If posDX > 0 Then
        temp = Left(temp, posDX - 1) & Mid(temp, endDX + 9)
    End If
    
    'removes {dx} tags, if present
    Dim posD, endD As Integer
    posD = InStr(temp, "{dx")
    endD = InStr(temp, "{/dx")
    If posD > 0 Then
        temp = Left(temp, posD - 1) & Mid(temp, endD + 5)
    End If
        
    'replaces "\u2013" with the appropriate "–"
    Dim posDash As Integer
    posDash = InStr(temp, "\u2013")
    While posDash > 1
        temp = Left(temp, posDash - 1) & "–" & Mid(temp, posDash + 6)
        posDash = InStr(temp, "\u2013")
    Wend
    
    getTransformedDef = temp

End Function

'when a new word is added to the list, this sub is called; the word's part of speech is printed in the cell to its immediate right
Sub printPoS(row As Integer, col As Integer)
    
    Cells(row, col).Value = getPoS(Cells(row, col - 1).Value)
    
End Sub

'when a new word is added to the list, this sub is called; the word's definition is printed in the cell two places to its right
Sub printDef(row As Integer, col As Integer)
    
    Cells(row, col).Value = getTransformedDef(Cells(row, col - 2).Value)

End Sub

'when a new word is added to the list, this sub is called; a link to the word's pronunciation is placed in the cell three places to its right
Sub printPronunciation(row As Integer, col As Integer)
        
    ActiveSheet.Hyperlinks.Add Cells(row, col), Address:=getPronunciation(Cells(row, col - 3).Value), TextToDisplay:=ChrW(&HD83D) & ChrW(&HDD0A)
'    Call stylizeHyperlinks

End Sub

'transforms input word into its default id
Sub transformInput(row As Integer, col As Integer)

    Cells(row, col).Value = getID(Cells(row, col).Value)

End Sub

'debug function: output that assesses the nature of the json from the api call given a specified word
Function outputRawJSon() As String
    
    Dim Json As Object
    Set Json = JsonConverter.ParseJson(getData(getURL("caught")))
    
   ' Debug.Print (Len(JsonConverter.ConvertToJson(Json))) 'JsonConverter.ConvertToJson(Json, Whitespace:=2))
    Debug.Print (JsonConverter.ConvertToJson(Json(1), Whitespace:=2))
    'JsonConverter.ConvertToJson (Json(1)("def")(1)("sseq")(1)(1)(2)("dt")(1)(2))
End Function

'debug function: clears the immediate window
Sub clearImmediate()
    
    Debug.Print Now
    Application.SendKeys "^g ^a {DEL}"

End Sub



