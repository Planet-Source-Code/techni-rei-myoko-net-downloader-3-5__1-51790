Attribute VB_Name = "INIreadwrite"
Option Explicit
Public Function issection(value As String) As Boolean
On Error Resume Next
    If Left(value, 1) = "[" And Right(value, 1) = "]" And stripsection(value) <> Empty Then issection = True Else issection = False
End Function
Public Function isvalue(value As String) As Boolean
On Error Resume Next
    If issection(value) = False And InStr(value, "=") > 0 Then isvalue = True Else isvalue = False
End Function
Public Function stripsection(section As String) As String
On Error Resume Next
    stripsection = Mid(section, 2, Len(section) - 2)
End Function
Public Function stripvalue(value As String) As String
On Error Resume Next
    stripvalue = Right(value, Len(value) - InStr(value, "="))
End Function
Public Function stripname(value As String) As String
On Error Resume Next
    stripname = Left(value, InStr(value, "=") - 1)
End Function
Public Function iscomment(value As String) As Boolean
On Error Resume Next
    If Left(value, 1) = "#" Or Left(value, 1) = "'" Then iscomment = True Else iscomment = False
End Function
Public Function getvalue(Filename As String, section As String, value As String, Optional default As String = Empty) As String
    On Error Resume Next
    getvalue = default
    Dim tempfile As Long, found As Boolean, temp As String, currentsection As String
    If fileexists(Filename) = True Then
        tempfile = FreeFile
        Open Filename For Input As #tempfile
            Do Until EOF(tempfile) Or found = True
                Line Input #tempfile, temp
                If iscomment(temp) = False Then
                    If issection(temp) = True Then currentsection = stripsection(temp)
                    If LCase(currentsection) = LCase(section) And isvalue(temp) = True Then
                        If LCase(stripname(temp)) = LCase(value) Then
                            getvalue = stripvalue(temp)
                            found = True
                        End If
                    End If
                End If
            Loop
        Close #tempfile
    End If
End Function
