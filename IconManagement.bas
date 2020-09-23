Attribute VB_Name = "IconManagement"
Option Explicit

Const uniqueicons As String = "*.scr;*.exe;*.ico;*.lnk;*.cpl;*.msc"

Public Function islike(Filter As String, expression As String) As Boolean
On Error Resume Next
Dim tempstr() As String, count As Long
If Replace(Filter, ";", Empty) <> Filter Then
tempstr = Split(Filter, ";")
islike = False
For count = LBound(tempstr) To UBound(tempstr)
    If LCase(expression) Like LCase(tempstr(count)) Then islike = True
Next
Else
If expression Like Filter Then islike = True Else islike = False
End If
End Function
Public Function isadir(Filename As String) As Boolean
On Error Resume Next
If Filename <> Empty Then If (GetAttr(Filename) And vbDirectory) = vbDirectory Then If InStr(Filename, "\") > 0 Then isadir = True
End Function
Public Function geticon(ByVal Filename As String, iml As ImageList, picture As PictureBox) As Long
    Dim count As Long
    If islike(uniqueicons, Filename) Then 'Or (isadir(filename) And FileExists(chkdir(filename, "desktop.ini"))) Then
        'is a file type with a unique icon, or is a folder with a unique icon. search by full filename
        count = searchicon(Filename, iml)
    Else
        If isadir(Filename) = True Then Filename = ".Folder"  'is a normal folder
        'search by extention
        Filename = Right(Filename, Len(Filename) - InStrRev(Filename, ".") + 1)
        count = searchicon(Filename, iml)
    End If
    
    If count = 0 Then
        geticon = createicon(Filename, iml, picture)
    Else
        geticon = count
    End If
End Function

Public Function searchicon(Filename As String, iml As ImageList) As Long
searchicon = 0
Dim count As Long
Filename = LCase(Filename)
For count = 1 To iml.ListImages.count
     If Filename = iml.ListImages.Item(count).tag Then
          searchicon = count
          Exit For
     End If
Next
End Function

Public Function createicon(ByVal Filename As String, iml As ImageList, picture As PictureBox) As Long
     Dim count As Long
     picture.Cls
     picture.Width = iml.ImageWidth * 15
     picture.Height = iml.ImageHeight * 15
     If InStr(Filename, ",") = 0 Then
        drawfileicon Filename, IIf(iml.ImageHeight = 16, SmallIcon, largeIcon), picture.hDC, 0, 0
     Else
        drawfileicon Left(Filename, InStrRev(Filename, ",") - 1), IIf(iml.ImageHeight = 16, SmallIcon, largeIcon), picture.hDC, 0, 0 ', Val(Right(Filename, Len(Filename) - InStrRev(Filename, ",")))
     End If
     count = iml.ListImages.count
     iml.ListImages.add , , picture.Image
     Do Until iml.ListImages.count > count
        DoEvents
     Loop
     iml.ListImages.Item(count + 1).tag = LCase(Filename)
     createicon = count + 1
End Function
