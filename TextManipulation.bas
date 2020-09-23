Attribute VB_Name = "TextManipulation"
Option Explicit
Public Const tags As String = "a" 'add any tag that needs a '</' & tagname & '>' to end it
Public Const imgs As String = "bmp gif jpg jpeg jpe jfif png table tr td"
Public Function chkurl(ByVal BASEHREF As String, url As String) As String
'check for absolute (is like *://*)
'check for relative (contains ../)
'check for additive (else)
Dim spoth As Long
If Left(url, 1) = "#" Then Exit Function 'is not a file
If Left(url, 1) = "/" Then url = Right(url, Len(url) - 1)
'If Len(basehref) - Len(Replace(basehref, "/", Empty)) = 2 Then basehref = basehref & "/"
If containsword(BASEHREF, "://") = False Then BASEHREF = "http://" & BASEHREF
If LCase(url) <> LCase(BASEHREF) And url <> Empty And BASEHREF <> Empty Then
If url Like "*://*" Then 'is absolute
    chkurl = url
Else
    If containsword(url, "../") Then 'is relative
        If Right(BASEHREF, 1) = "/" And Len(BASEHREF) - Len(Replace(BASEHREF, "/", Empty)) > 2 Then BASEHREF = Left(BASEHREF, Len(BASEHREF) - 1)
        If containsword(Replace(BASEHREF, "://", ""), "/") = True Then
            For spoth = 1 To countwords(BASEHREF, "../")
                url = Right(url, Len(url) - Len("../"))
                BASEHREF = Left(BASEHREF, InStrRev(BASEHREF, "/"))
            Next
        Else
            url = Replace(url, "../", "")
        End If
        If Right(BASEHREF, 1) <> "/" Then chkurl = BASEHREF & "/" & url Else chkurl = BASEHREF & url
    Else 'is additive
        If Right(BASEHREF, 1) <> "/" Then chkurl = BASEHREF & "/" & url Else chkurl = BASEHREF & url
    End If
End If
End If
End Function
Public Function containsword(text As String, word As String) As Boolean
    containsword = InStr(1, text, word, vbTextCompare) > 0
End Function
Public Function countwords(text As String, word As String) As Long
    countwords = (Len(text) - Len(Replace(text, word, Empty, , , vbTextCompare))) \ Len(word)
End Function

Public Function enumHTMLTAGS(ByVal HTMLCode As String, stringarray, upbound As Long, TAGTYPE As String, ParamArray DESIREDPROPERTY() As Variant) As Long
    Dim temp As Long, tempstr As String, doop As Boolean, temparr() As String, temp2 As Long
    Do Until Len(HTMLCode) = 0
        Select Case Left(HTMLCode, 1)
            Case "<" 'is an html tag
                
                temp2 = InStr(1, HTMLCode, "</" & tagname(HTMLCode) & ">", vbTextCompare)
                
                If temp2 > 0 And itemexists(tags, " ", tagname(HTMLCode)) Then 'Not itemexists(tags, " ", tagname(HTMLCODE)) Then ' go to </ & tagname
                    tempstr = Left(HTMLCode, 2 + Len(tagname(HTMLCode)) + temp2)
                Else
                    tempstr = Left(HTMLCode, InStr(1, HTMLCode, ">", vbTextCompare))
                End If
                
                doop = False 'Do Operation
                If TAGTYPE = "*" Then doop = True
                If StrComp(TAGTYPE, tagname(tempstr), vbTextCompare) = 0 Then doop = True
                If LCase(TAGTYPE) = "hrefimg" And LCase(tagname(tempstr)) = "a" And InStr(1, tempstr, "<img", vbTextCompare) > 0 Then doop = True
                                
                'MsgBox HTMLCODE, , doop & " " & tempstr
                                
                If doop Then
                    ReDim temparr(0 To UBound(DESIREDPROPERTY))
                    For temp = 0 To UBound(DESIREDPROPERTY)
                        temparr(temp) = addfrom(tempstr, DESIREDPROPERTY(temp) & Empty)
                    Next
                    enumHTMLTAGS = append(stringarray, upbound, temparr)
                End If
                
                HTMLCode = Right(HTMLCode, Len(HTMLCode) - Len(tempstr))
                
            Case Else 'isnt
                temp = InStr(HTMLCode, "<")
                If temp = 0 Then
                    HTMLCode = Empty
                Else
                    HTMLCode = Right(HTMLCode, Len(HTMLCode) - temp + 1)
                End If
        End Select
        DoEvents
    Loop
End Function

Public Function append(destination, upbound As Long, Items) As Long
    Dim higher As Long, count As Long 'handles two dimensional zero indexed first dimension, one indexed second dimension arrays only
    higher = UBound(Items)
    
    upbound = upbound + 1
    If upbound = 1 Then
        ReDim destination(0 To higher, 1 To 1)
    Else
        ReDim Preserve destination(0 To higher, 1 To upbound)
    End If
    
    For count = 0 To higher
        destination(count, upbound) = Items(count) & Empty
    Next
    
    append = higher
End Function
Public Function addfrom(content As String, tag As String) As String
    Dim temp As Long, location As Long, temp2 As Long
    If LCase(tag) <> "node" Then
    
    location = InStr(1, content, tag, vbTextCompare)
    If location > 0 Then
    location = InStr(location, content, "=") + 1
    Select Case Mid(content, location, 1)
        Case """", "'"
            location = location + 1
            temp = InStr(location, content, """")
            If temp = 0 Then temp = InStr(location, content, "'")
            temp2 = InStr(location, content, ">")
        Case Else
            temp = InStr(location, content, " ")
            temp2 = InStr(location, content, ">")
    End Select
    If temp2 < temp And temp2 > 0 Then temp = temp2
    If temp = 0 Then temp = InStr(location, content, ">")
    If temp = 0 Then temp = Len(content)
    addfrom = Mid(content, location, temp - location)
    End If
    
    Else
    
    'temp = InStr(content, ">")
    'temp2 = InStrRev(content, "<")
    'addfrom = Mid(content, temp + 1, temp2 - temp - 1)
    addfrom = removebrackets(content, "<", ">")
    
    End If
End Function
Public Function hastag(content As String, tag As String) As Boolean
    hastag = InStr(1, content, tag & "=", vbTextCompare) > 0
End Function
Public Function tagname(content As String) As String
    Dim temp As Long, temp2 As Long
    temp = InStr(content, " ")
    temp2 = InStr(content, ">")
    If temp > 0 And temp < temp2 Then temp2 = temp
    tagname = Mid(content, 2, temp2 - 2)
End Function
Public Function itemexists(list As String, delimeter As String, ByVal name As String) As Boolean
    On Error Resume Next
    If InStr(list, delimeter) > 0 Then
        Dim flist() As String
        flist = Filter(Split(list, delimeter), name, , vbTextCompare)
        itemexists = flist(0) = name
    Else
        itemexists = StrComp(list, name, vbTextCompare) = 0
    End If
End Function
Public Function makeabsolute(text As String) As String
    makeabsolute = text
    If InStr(text, "/") > 0 Then makeabsolute = Right(text, Len(text) - InStrRev(text, "/"))
End Function

Public Function isanimage(ByVal Filename As String) As Boolean
    If InStr(Filename, ".") > 0 Then
        Filename = Right(Filename, Len(Filename) - InStrRev(Filename, "."))
        isanimage = itemexists(imgs, " ", Filename)
    End If
End Function
Public Function removetext(text As String, start As Long, finish As Long, Optional exclusive As Boolean = True) As String
    If exclusive = True Then
        removetext = Left(text, start - 1) & Right(text, Len(text) - finish)
    Else
        removetext = Mid(text, start, finish - start)
    End If
End Function
Public Function removebrackets(ByVal text As String, leftb As String, rightb As String) As String
    Do While InStr(text, leftb) > 0 And InStr(text, rightb) > InStr(text, leftb)
        text = removetext(text, InStr(text, leftb), InStr(text, rightb))
    Loop
    removebrackets = text
End Function

Public Function GetExtention(url As String) As String
    Dim temp As Long, temp2 As Long, temp3 As Long, tempstr As String
    tempstr = ".html"
    temp = InStrRev(url, ".")
    temp2 = InStrRev(url, "/")
    temp3 = InStr(url, "://")
    
    If temp > temp2 And temp2 > temp3 + 2 Then
        tempstr = "." & Right(url, Len(url) - temp)
        temp = InStr(tempstr, "?")
        If temp > 0 Then tempstr = Left(tempstr, temp - 1)
    End If
    If InStrRev(tempstr, ".") > 1 Then tempstr = Right(tempstr, Len(tempstr) - InStrRev(tempstr, "."))
    If InStr(tempstr, "?") > 1 Then tempstr = Right(tempstr, Len(tempstr) - InStr(tempstr, "?"))
    GetExtention = tempstr
End Function
Public Function killallceptnumber(text As String, Optional delimeter As String = " ", Optional reverse As Boolean) As String
    Dim temp As Long, tempstr As String
    For temp = 1 To Len(text)
        Select Case Mid(text, temp, 1)
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                If Not reverse Then
                    tempstr = tempstr & Mid(text, temp, 1)
                Else
                    If Len(tempstr) > 0 Then If Right(tempstr, 1) <> delimeter Then tempstr = tempstr & delimeter
                End If
            Case Else
                If Not reverse Then
                    If Len(tempstr) > 0 Then If Right(tempstr, 1) <> delimeter Then tempstr = tempstr & delimeter
                Else
                    tempstr = tempstr & Mid(text, temp, 1)
                End If
        End Select
    Next
    If tempstr = Empty Then tempstr = -1
    killallceptnumber = tempstr
End Function
Public Function num2patt(text As String) As String
    Dim temp As Long, tempstr As String
    For temp = 1 To Len(text)
        'If StrComp(Mid(text, temp, 1), Mid(text2, temp, 1), vbTextCompare) <> 0 Then
            Select Case Mid(text, temp, 1)
                Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9": tempstr = tempstr & "#"
                Case Else: tempstr = tempstr & Mid(text, temp, 1)
            End Select
        'Else
        '    tempstr = tempstr & Mid(text, temp, 1)
        'End If
    Next
    num2patt = tempstr
End Function
Public Function isvalidpattern(text As String, Optional del As String = "#") As Boolean
    Dim temp As Long, foundstart As Boolean, foundend As Boolean, buffer As Boolean
    buffer = True
    For temp = 1 To Len(text)
        If Mid(text, temp, 1) = del Then
            foundstart = True
            If foundend = True Then buffer = False
        Else
            If foundstart = True Then foundend = True
        End If
    Next
    isvalidpattern = buffer
End Function

Public Function seedstring(text As String, ByVal instring As String, Optional del As String = "#", Optional default As String = "0")
    Dim temp As Long, tempstr As String
    For temp = Len(text) To 1 Step -1
        If Mid(text, temp, 1) = del Then
            If Len(instring) > 0 Then
                tempstr = tempstr & Right(instring, 1)
                instring = Left(instring, Len(instring) - 1)
            Else
                tempstr = tempstr & default
            End If
        Else
            tempstr = tempstr & Mid(text, temp, 1)
        End If
    Next
    seedstring = StrReverse(tempstr)
End Function
Public Function isCancel(ByVal tag As String) As Boolean
    If Left(tag, 1) = "<" Then tag = Right(tag, Len(tag) - 1)
    isCancel = Left(tag, 1) = "/"
End Function
Public Function CleanTag(ByVal tag As String) As String
    tag = tagname(tag)
    If Left(tag, 1) = "/" Then tag = Right(tag, Len(tag) - 1)
    CleanTag = LCase(tag)
End Function
