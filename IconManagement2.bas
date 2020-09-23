Attribute VB_Name = "IconManagement2"
Option Explicit
'icon sizelocated in GetString(HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics", "Shell Icon Size", 32)
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As typSHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal Flags&) As Long

Public Type typSHFILEINFO
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * 260
  szTypeName As String * 80
End Type

Public Enum IconSize
    LargeIcon = 0
    SmallIcon = 1
End Enum

Public Const SH_USEFILEATTRIBUTES As Long = &H10
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const SHGFI_DISPLAYNAME  As Long = &H200
Public Const SHGFI_EXETYPE  As Long = &H2000
Public Const SHGFI_SYSICONINDEX  As Long = &H4000
Public Const SHGFI_SHELLICONSIZE  As Long = &H4
Public Const SHGFI_TYPENAME  As Long = &H400
Public Const SHGFI_LARGEICON  As Long = &H0
Public Const SHGFI_SMALLICON  As Long = &H1
Public Const ILD_TRANSPARENT As Long = &H1
Public Const Flags = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE Or SH_USEFILEATTRIBUTES
Public FileInfo As typSHFILEINFO

Public Function geticonhandle(Filename As String, Size As Long) 'Gets a handle to the icon
    geticonhandle = SHGetFileInfo(Filename, FILE_ATTRIBUTE_NORMAL, FileInfo, Len(FileInfo), Flags Or Size)
End Function

Public Function drawfileicon(filetype As String, Size As IconSize, destHDC As Long, x As Long, y As Long) 'Draws the icon int the destination.hdc
    drawfileicon = ImageList_Draw(geticonhandle(filetype, Size), FileInfo.iIcon, destHDC, x, y, ILD_TRANSPARENT)
End Function

Public Function HasUniqueIcon(Filename As String) As Boolean
    HasUniqueIcon = GetDefaultIcon(GetClassname(GetExtention(Filename))) = "%1"
End Function

Public Function GetFilenoext(ByVal Filename As String) As String
    Dim temp As Long
    temp = InStrRev(Filename, "\")
    If temp > 0 Then Filename = Right(Filename, Len(Filename) - temp)
    temp = InStrRev(Filename, ".")
    If temp > 0 Then Filename = Left(Filename, temp - 1)
    GetFilenoext = Filename
End Function

Public Function GetFilename(Filename As String) As String
    Dim temp As Long
    temp = InStrRev(Filename, "\")
    If temp = 0 Then
        GetFilename = Filename
    Else
        GetFilename = Right(Filename, Len(Filename) - temp)
    End If
End Function

Public Function GetPath(Filename As String) As String
    If InStr(Filename, "\") > 0 Then GetPath = Left(Filename, InStrRev(Filename, "\") - 1) Else GetPath = Filename
End Function

Public Function GetIconSize(Optional Size As IconSize = LargeIcon) As Long
    If Size = LargeIcon Then GetIconSize = CLng(GetString(HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics", "Shell Icon Size", 32)) Else GetIconSize = 16
End Function

Public Function GetClassname(ByVal Extention As String) As String
    If Left(Extention, 1) <> "." Then Extention = "." & Extention
    GetClassname = GetString(HKEY_CLASSES_ROOT, Extention)
End Function

Public Function GetDefaultIcon(Classname As String) As String
    GetDefaultIcon = GetString(HKEY_CLASSES_ROOT, Classname & "\DefaultIcon")
End Function

Public Function IsADir(Filename As String) As Boolean
    On Error Resume Next
    If Len(Filename) > 0 Then IsADir = (GetAttr(Filename) And vbDirectory) = vbDirectory
End Function

Public Function IsLike(ByVal text As String, ByVal Expression As String) As Boolean
    Dim temp As Long, tempstr() As String
    text = LCase(text)
    Expression = LCase(Expression)
    If InStr(Expression, ";") = 0 Then
        IsLike = text Like Expression
    Else
        tempstr = Split(Expression, ";")
        For temp = 0 To UBound(tempstr)
            If text Like tempstr(temp) Then
                IsLike = True
                Exit Function
            End If
        Next
    End If
End Function

Public Function GetIndex(Key As String, IML As ImageList) As Long
    On Error Resume Next
    GetIndex = IML.ListImages.item(Key).Index
End Function

Public Function GetIcon(ByVal Filename As String, IML As ImageList, picture As PictureBox) As Long
    Dim count As Long, OldFilename As String
    Filename = Trim(Filename)
    OldFilename = Filename
    count = GetIndex(Filename, IML)
    
    If count = 0 Then
        GetIcon = CreateIcon(OldFilename, Filename, IML, picture)
    Else
        GetIcon = count
    End If
End Function

Public Function CreateIcon(Filename As String, Key As String, IML As ImageList, picture As PictureBox, Optional Size As IconSize = SmallIcon) As Long
     Dim count As Long
     picture.Cls
     picture.ScaleHeight = GetIconSize(Size)
     picture.ScaleWidth = picture.ScaleHeight
     drawfileicon Filename, Size, picture.hDC, 0, 0
     count = IML.ListImages.count
     IML.ListImages.Add , Key, picture.Image
     CreateIcon = count + 1
End Function

Public Function chkdir(Path As String, Filename As String) As String
    chkdir = Path & IIf(Right(Path, 1) = "\", Empty, "\") & Filename
End Function
