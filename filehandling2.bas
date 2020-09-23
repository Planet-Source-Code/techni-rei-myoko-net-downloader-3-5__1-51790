Attribute VB_Name = "filehandling"
Option Explicit

Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFilename As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Public Function uniquefilename(FileName As String) As String
    Dim temp1 As String, temp2 As String, temp3 As Long
    uniquefilename = FileName
    
    If FileExists(FileName) Then
        Dim count As Long
        count = 1
        temp3 = InStrRev(FileName, ".")
        temp1 = FileName
        If temp3 > 0 Then
            temp1 = left(FileName, temp3 - 1)
            temp2 = Right(FileName, Len(FileName) - temp3 + 1)
        End If
        Do Until FileExists(temp1 & " (" & count & ")" & temp2) = False
            count = count + 1
        Loop
        uniquefilename = temp1 & " (" & count & ")" & temp2
    End If
End Function
Public Function DownloadFile(url As String, FileName As String) As Boolean
On Error Resume Next
If Len(FileName) > 255 Then FileName = left(FileName, 255)

FileName = Replace(FileName, "*", Empty)
FileName = Replace(FileName, "&", Empty)
FileName = Replace(FileName, "%", Empty)
FileName = Replace(FileName, "=", Empty)

If InStrRev(FileName, "?") > 0 Then
    Dim temp As String
    temp = left(FileName, InStrRev(FileName, "?") - 1)
    If InStr(temp, ".") = 0 Then temp = temp & ".txt"
    FileName = uniquefilename(temp)
End If

'Downloads the file from URL and saves it as filename
If url <> Empty And FileName <> Empty Then DownloadFile = URLDownloadToFile(0, url, FileName, 0, 0) = 0
DoEvents
End Function

'I could make these 2 into one function, but I like it seperate
Public Function direxists(directory As String) As Boolean
'Checks to see if a directory exists
On Error Resume Next
If Dir(directory, vbDirectory + vbHidden) = Empty Then direxists = False Else direxists = True
End Function
Public Function FileExists(FileName As String) As Boolean
'Checks to see if a file exists
On Error Resume Next
If Dir(FileName) = Empty Then FileExists = False Else FileExists = True
End Function

Public Function chkfile(directory As String, FileName As String) As String
'Adds the filename to a dir without getting an error if its the root dir
On Error Resume Next
If Right(directory, 1) <> "\" Then chkfile = directory & "\" & FileName Else chkfile = directory & FileName
End Function

Public Function loadfile(FileName As String) As String
On Error Resume Next
Dim intFile As Integer, temp As String, allfile As String
allfile = Empty
If Dir(FileName) <> Empty And Right(FileName, 1) <> "\" And FileLen(FileName) > 0 Then
intFile = FreeFile()
Open FileName For Input As intFile
Do Until EOF(intFile)
    Line Input #intFile, temp
    allfile = allfile & temp & vbNewLine
Loop
Close intFile
loadfile = left(allfile, Len(allfile) - 1)
Else
loadfile = Empty
End If
End Function
