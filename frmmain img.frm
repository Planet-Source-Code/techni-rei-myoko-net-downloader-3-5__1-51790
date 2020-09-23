VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   Caption         =   "Net Downloader"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   7065
   Icon            =   "frmmain img.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin NetDownloader.TrillianFrame trlmain 
      Height          =   3855
      Index           =   2
      Left            =   120
      TabIndex        =   22
      Top             =   3120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   6800
      Caption         =   "Referenced files"
      Begin VB.CheckBox chkall 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Set all selected to me"
         Height          =   255
         Left            =   4920
         TabIndex        =   34
         Top             =   420
         Width           =   1815
      End
      Begin VB.CommandButton cmdmain 
         Caption         =   "Stop"
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   32
         Top             =   3000
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdmain 
         Caption         =   "Clear all Links"
         Height          =   375
         Index           =   7
         Left            =   4080
         TabIndex        =   31
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton cmdmain 
         Caption         =   "Clear"
         Height          =   375
         Index           =   3
         Left            =   5400
         TabIndex        =   30
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton cmdmain 
         Caption         =   "DL All Links"
         Height          =   375
         Index           =   6
         Left            =   1440
         TabIndex        =   29
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton cmdmain 
         Caption         =   "DL ImageLinks"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   28
         Top             =   3360
         Width           =   1335
      End
      Begin VB.PictureBox picmain 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   5640
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   27
         Top             =   1080
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.CommandButton cmdmain 
         Caption         =   "View All Links"
         Height          =   375
         Index           =   5
         Left            =   1920
         TabIndex        =   24
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdmain 
         Caption         =   "View Image Links"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1815
      End
      Begin MSComctlLib.ListView lstmain 
         Height          =   2655
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   4683
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "URL"
            Object.Width           =   176
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Image"
            Object.Width           =   176
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Base Href"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstsec 
         Height          =   2655
         Left            =   3720
         TabIndex        =   25
         Top             =   720
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imlico"
         SmallIcons      =   "imlico"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Comment"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "URL"
            Object.Width           =   2646
         EndProperty
      End
      Begin MSComctlLib.ImageList imlthumbs 
         Left            =   5400
         Top             =   840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   16777215
         MaskColor       =   16777215
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList imlico 
         Left            =   6000
         Top             =   840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmmain img.frx":0E42
               Key             =   ".php"
               Object.Tag             =   ".php"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmmain img.frx":1194
               Key             =   ".htm"
               Object.Tag             =   ".htm"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmmain img.frx":14E6
               Key             =   ".html"
               Object.Tag             =   ".html"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmmain img.frx":1838
               Key             =   ".shtml"
               Object.Tag             =   ".shtml"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmmain img.frx":1B8A
               Key             =   ".cgi"
               Object.Tag             =   ".cgi"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmmain img.frx":1EDC
               Key             =   ".eml"
               Object.Tag             =   ".eml"
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdmain 
         Caption         =   "Recurse All Links"
         Height          =   375
         Index           =   9
         Left            =   2760
         TabIndex        =   33
         Top             =   3360
         Width           =   1335
      End
   End
   Begin NetDownloader.TrillianFrame trlmain 
      Height          =   1335
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   2355
      Caption         =   "Seed with artificial links"
      Begin VB.CommandButton cmdadd 
         Caption         =   "&Add"
         Height          =   285
         Left            =   6000
         TabIndex        =   19
         ToolTipText     =   "Add a range of files to the list from the URL typed above"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtubound 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5400
         TabIndex        =   18
         Text            =   "100"
         ToolTipText     =   "Put the highest number here"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtlbound 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4440
         TabIndex        =   16
         Text            =   "0"
         ToolTipText     =   "Put the lowest number here (Usually 1 or 0)"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtfiles 
         Height          =   285
         Left            =   1080
         TabIndex        =   13
         ToolTipText     =   "Type the name of the file range here, replacing the numbers with #. If it's numbered like 001, then replace it with ###"
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label lblmain 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "# signs will be replaced with numbers from Lo to Hi, padded to fit the number of # signs"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   6615
      End
      Begin VB.Image imgmain 
         Height          =   360
         Left            =   120
         Picture         =   "frmmain img.frx":222E
         ToolTipText     =   "Extract the required data from an EZCode"
         Top             =   900
         Width           =   915
      End
      Begin VB.Label lblmain 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Example: page###.txt for 39 becomes page039.txt"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   6615
      End
      Begin VB.Label lblmain 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hi:"
         Height          =   255
         Index           =   5
         Left            =   5040
         TabIndex        =   17
         Top             =   390
         Width           =   255
      End
      Begin VB.Label lblmain 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lo:"
         Height          =   255
         Index           =   4
         Left            =   4080
         TabIndex        =   15
         Top             =   390
         Width           =   255
      End
      Begin VB.Label lblmain 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pattern:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   390
         Width           =   855
      End
   End
   Begin NetDownloader.TrillianFrame trlmain 
      Height          =   1515
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   2672
      Caption         =   "Download details/parameters"
      Begin VB.CheckBox chksec 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Email links"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   6855
         TabIndex        =   35
         Tag             =   "/e"
         ToolTipText     =   "Requires normal links to be checked"
         Top             =   1080
         Width           =   1320
      End
      Begin VB.CheckBox chksec 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Scripts"
         Height          =   255
         Index           =   3
         Left            =   5400
         TabIndex        =   10
         Tag             =   "/s"
         Top             =   1080
         Width           =   840
      End
      Begin VB.CheckBox chksec 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Backgrounds"
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   9
         Tag             =   "/b"
         Top             =   1080
         Width           =   1320
      End
      Begin VB.CheckBox chksec 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Images"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   8
         Tag             =   "/i"
         Top             =   1080
         Width           =   840
      End
      Begin VB.CheckBox chksec 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Links"
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   7
         Tag             =   "/l"
         Top             =   1080
         Width           =   720
      End
      Begin VB.ComboBox txtmain 
         Height          =   315
         Index           =   1
         Left            =   1080
         TabIndex        =   6
         Top             =   720
         Width           =   4815
      End
      Begin VB.CommandButton cmdmain 
         Caption         =   "&Browse"
         Height          =   315
         Index           =   1
         Left            =   6000
         TabIndex        =   4
         Top             =   720
         Width           =   735
      End
      Begin VB.ComboBox txtmain 
         Height          =   315
         Index           =   0
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   4815
      End
      Begin VB.CommandButton cmdmain 
         Caption         =   "&Go"
         Height          =   315
         Index           =   0
         Left            =   6000
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblmain 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search for:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblmain 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Destination"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   750
         Width           =   855
      End
      Begin VB.Label lblmain 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Source:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   390
         Width           =   855
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "file"
      Visible         =   0   'False
      Begin VB.Menu mnucheck 
         Caption         =   "Check Selected"
         Index           =   0
      End
      Begin VB.Menu mnucheck 
         Caption         =   "Uncheck Selected"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim supress As Boolean, isclicking As Boolean

Private Sub chkall_Click()
    mnucheck_Click 1 - chkall.value
End Sub

Private Sub chkall_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then chkall_Click
End Sub

Private Sub chksec_Click(Index As Integer)
If Index = 2 Then chksec(4).Enabled = chksec(2).value = vbChecked
End Sub

Private Sub cmdadd_Click()
Dim spoth As Long, spoth2 As Long
'If Not isvalidpattern(txtfiles) Then
'    MsgBox "This is not a valid pattern, all of the #'s must be put together", vbCritical, "Not a valid pattern"
'Else
If txtmain(0) <> Empty Then
spoth2 = countwords(txtfiles, "#")
For spoth = Val(txtlbound) To Val(txtubound) 'countchars -1
    With lstsec.ListItems
        .Add , , "Added by user (" & spoth & ")"
        .item(.count).Checked = True
        If spoth2 = 1 Then
            .item(.count).SubItems(1) = chkurl(txtmain(0), Replace(txtfiles, "#", spoth))
        Else
            '.item(.count).SubItems(1) = chkurl(txtmain(0), Replace(txtfiles, String(spoth2, "#"), Format(spoth, String(spoth2, "0"))))
            .item(.count).SubItems(1) = chkurl(txtmain(0), seedstring(txtfiles, spoth))
        End If
        .item(.count).tag = GetIcon("." & Right(txtfiles, Len(txtfiles) - InStrRev(txtfiles, ".")), Me.imlico, Me.picmain)
        If imlico.ListImages.count = 1 Then Set lstsec.SmallIcons = imlico
        .item(.count).SmallIcon = Val(.item(.count).tag)
    End With
    DoEvents
Next
removedoubles lstsec, 2
resizecolumnheaders lstsec
Else
MsgBox "I need a site to download the files from first", vbCritical, "No URL given."
End If
'End If
End Sub

Public Sub cmdmain_Click(Index As Integer)
    If Not direxists(txtmain(1)) Then MkDir txtmain(1)
    If Not direxists(txtmain(1)) Then
        MsgBox "Unable to create non-existant folder. Aborting.", vbCritical, "Unable to complete operation"
        Exit Sub
    End If
    Select Case Index
        Case 0
            SaveSetting "Net Downloader", "Main", "LastURL", txtmain(0).text
            lstsec.Sorted = False
            cmdmain(8).Visible = True
            enumimagelinks txtmain(0) 'URL
            cmdmain(8).Visible = False
            If imlthumbs.ListImages.count = 0 Then
                cmdmain_Click 3
                If isclicking Then lstsec.selecteditem.ForeColor = vbRed
            End If
        Case 1 'DIR
            'BB.InitDir = txtmain(1)
            'BB.ShowButton = True
            'txtmain(1).tag = BrowseFF
            txtmain(1).tag = BrowseForFolder(Me.hwnd, "Please select a folder")
            If Len(txtmain(1).tag) > 0 Then
                txtmain(1) = txtmain(1).tag
                If Not itemexists(txtmain(1), txtmain(1).tag) Then txtmain(1).additem txtmain(1).tag
            End If
        Case 2 'download and clear
            cmdmain(8).Visible = True
            downloadselected
            cmdmain(8).Visible = False
            If Not supress Then MsgBox txtmain(0) & " was extracted succesfully", vbInformation, "Operation complete"
            cmdmain_Click 3
        Case 3 'clear
            If cmdmain(8).Visible Then cmdmain_Click 8
            DoEvents
            lstmain.ListItems.Clear
            Set lstmain.icons = Nothing
            Set lstmain.SmallIcons = Nothing
            imlthumbs.ListImages.Clear
            txtmain(0) = Empty
            If lstsec.ListItems.count > 0 Then cmdmain_Click 5
        Case 4
            lstmain.Visible = True
            cmdmain(4).Font.Bold = True
            cmdmain(5).Font.Bold = False
        Case 5
            lstmain.Visible = False
            cmdmain(4).Font.Bold = False
            cmdmain(5).Font.Bold = True
        Case 6 'download links
            downloadlinks
            If Not supress Then MsgBox txtmain(0) & " was extracted succesfully", vbInformation, "Operation complete"
            cmdmain_Click 3
        Case 7 'clear links
            lstsec.ListItems.Clear
        Case 8: cmdmain(8).Visible = False
        Case 9 'recurse all
            supress = True
            downloadall
            supress = False
            MsgBox "All links were extracted succesfully", vbInformation, "Operation complete"
    End Select
    Me.Caption = "Net Downloader"
End Sub
Public Sub downloadall()
    Dim temp As Long, temp2 As Long
    temp2 = lstsec.ListItems.count
    For temp = 1 To temp2
        If lstsec.ListItems(temp).Checked = True Then
            'getitem lstsec.ListItems(temp).SmallIcon
            Set lstsec.selecteditem = lstsec.ListItems(temp)
            lstsec_DblClick
        End If
    Next
End Sub
Public Function itemexists(list As Object, itemname As String) As Boolean
    Dim temp As Long
    For temp = 0 To list.ListCount - 1
        If StrComp(list.list(temp), itemname, vbTextCompare) = 0 Then
            itemexists = True
            Exit For
        End If
        DoEvents
    Next
End Function
Public Sub downloadlinks()
    Dim temp As Long, Filename As String, url As String
    Me.Caption = "Downloading Links"
    For temp = 1 To lstsec.ListItems.count
        With lstsec.ListItems.item(temp)
            If .Checked Then
                url = .SubItems(1)
                Filename = Right(url, Len(url) - InStrRev(url, "/"))
                Filename = uniquefilename(chkfile(txtmain(1), Filename))
                DownloadFile url, Filename
            End If
        End With
        Me.Caption = Round(temp / lstsec.ListItems.count * 100) & "% done downloading"
        DoEvents
    Next
End Sub
Public Sub downloadselected()
    Dim temp As Long, url As String, img As String
    Me.Caption = "Downloading Files"
    For temp = 1 To lstmain.ListItems.count
        If Not cmdmain(8).Visible Then Exit For
        With lstmain.ListItems.item(temp)
            If .Checked Then
                url = chkurl(.SubItems(2), .text)
                img = uniquefilename(chkfile(txtmain(1), makeabsolute(.text)))
                DownloadFile url, img
            End If
        End With
        Me.Caption = Round(temp / lstmain.ListItems.count * 100) & "% done downloading"
        DoEvents
    Next
End Sub
Private Sub Form_Load()
    txtmain(1) = GetSetting("Net Downloader", "Main", "Last Path", App.Path)
    WindowState = GetSetting("Net Downloader", "Main", "WindowState", WindowState)
    Width = GetSetting("Net Downloader", "Main", "Width", Width)
    Height = GetSetting("Net Downloader", "Main", "Height", Height)
    Top = GetSetting("Net Downloader", "Main", "Top", Top)
    Left = GetSetting("Net Downloader", "Main", "Left", Left)
    
    Dim temp As Long, temp2 As Long, tempstr As String, favdir As String
    For temp = 0 To chksec.UBound
        chksec(temp).value = GetSetting("Net Downloader", "Main", chksec(temp).Caption, vbUnchecked)
    Next
    
    temp2 = Val(GetSetting("Net Downloader", "Destination", "Count", 0))
    For temp = 1 To temp2
        txtmain(1).additem GetSetting("Net Downloader", "Destination", "Dir:" & temp)
    Next
    
    favdir = GetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Favorites")
    For temp = 1 To 25
        tempstr = GetString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\TypedURLs", "url" & temp)
        If InStr(tempstr, "://") > 0 Then 'is a url
            txtmain(0).additem tempstr
        Else
            If StrComp(Right(Trim(tempstr), 4), ".url", vbTextCompare) = 0 Then 'is a favorite
                tempstr = chkfile(favdir, tempstr)
                txtmain(0).additem getvalue(tempstr, "InternetShortcut", "URL")
                If Len(txtmain(0).list(txtmain(0).ListCount - 1)) = 0 Then txtmain(0).RemoveItem txtmain(0).ListCount - 1
            End If
        End If
    Next
    txtmain(0).text = GetSetting("Net Downloader", "Main", "LastURL")
        
    If Len(Command) > 0 Then
        If InStr(Command, "?") > 0 And InStrRev(Command, "?") > InStr(Command, "?") Then
            imgmain_Click
        Else
            HandleCommand Command
        End If
    End If
End Sub

Public Sub HandleCommand(ByVal Command As String)
    Dim tempstr As String, temp As Long, temp2 As Long
    temp = InStr(Command, """")
    temp2 = InStrRev(Command, """")
    tempstr = Mid(Command, temp + 1, temp2 - temp - 1)
    Command = Left(Command, temp - 1) & Right(Command, Len(Command) - temp2)
    For temp = chksec.LBound To chksec.UBound
        chksec(temp).value = IIf(InStr(1, Command, chksec(temp).tag, vbTextCompare) > 0, vbChecked, vbUnchecked)
    Next
    txtmain(0).text = tempstr
    cmdmain_Click 0
End Sub

Private Sub Form_Resize()
    If Me.Width > 0 And Me.Height > 0 And Me.WindowState <> vbMinimized Then
        'Make sure it meets the minimum size requirements
        If Me.Width < 7215 Then Me.Width = 7215
        If Me.Height < 7500 Then Me.Height = 7500
        
        'Resize horizontally
        trlmain(0).Width = Me.Width - 330
        trlmain(1).Width = trlmain(0).Width
        trlmain(2).Width = trlmain(0).Width
        
        cmdmain(0).Left = trlmain(0).Width - 855
        cmdmain(1).Left = cmdmain(0).Left
        cmdadd.Left = cmdmain(0).Left
        
        txtmain(0).Width = trlmain(0).Width - 2040
        txtmain(1).Width = txtmain(0).Width
        
        txtfiles.Width = trlmain(0).Width - 3960
        txtlbound.Left = trlmain(0).Width - 2415
        txtubound.Left = trlmain(0).Width - 1455
        
        lblmain(4).Left = trlmain(0).Width - 2775
        lblmain(5).Left = trlmain(0).Width - 1815
        lblmain(6).Width = trlmain(0).Width - 240
        lblmain(7).Width = lblmain(6).Width
        
        lstmain.Width = lblmain(6).Width
        lstsec.Left = lstmain.Left
        lstsec.Width = lblmain(6).Width
        chkall.Left = lblmain(6).Width - chkall.Width
        
        cmdmain(3).Left = lstmain.Width + lstmain.Left - cmdmain(3).Width
        cmdmain(7).Left = cmdmain(3).Left - cmdmain(7).Width
        cmdmain(8).Width = lblmain(6).Width - cmdmain(7).Width
        cmdmain(8).Left = cmdmain(2).Left
        cmdmain(9).Left = (cmdmain(8).Width - cmdmain(9).Width) / 2 + cmdmain(2).Left
        
        'Resize vertically
        trlmain(2).Height = Me.Height - 3645
        cmdmain(2).Top = trlmain(2).Height - 495
        cmdmain(3).Top = cmdmain(2).Top
        cmdmain(6).Top = cmdmain(2).Top
        cmdmain(7).Top = cmdmain(2).Top
        cmdmain(8).Top = cmdmain(2).Top
        cmdmain(9).Top = cmdmain(2).Top
        
        lstmain.Height = trlmain(2).Height - 1200
        lstsec.Height = lstmain.Height
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveSetting("Net Downloader", "Main", "Last Path", txtmain(1))
    Call SaveSetting("Net Downloader", "Main", "WindowState", WindowState)
    WindowState = 0
    Call SaveSetting("Net Downloader", "Main", "Width", Width)
    Call SaveSetting("Net Downloader", "Main", "Height", Height)
    Call SaveSetting("Net Downloader", "Main", "Top", Top)
    Call SaveSetting("Net Downloader", "Main", "Left", Left)
    
    Dim temp As Long, temp2 As Long
    For temp = 0 To chksec.UBound
        SaveSetting "Net Downloader", "Main", chksec(temp).Caption, chksec(temp).value
    Next
    
    temp2 = txtmain(1).ListCount
    SaveSetting "Net Downloader", "Destination", "Count", temp2
    For temp = 0 To temp2 - 1
        SaveSetting "Net Downloader", "Destination", "Dir:" & temp + 1, txtmain(1).list(temp)
    Next
End Sub

Public Sub enumatag(ByVal HTMLCode As String, BASEHREF As String, tag As String, prop As String, prop2 As String, icon As String, Optional useextention As Boolean)
    Dim test() As String, up As Long, higher As Long, count As Long, tempstr As String, temp2 As Long, temp3 As Long
    higher = enumHTMLTAGS(HTMLCode, test, up, tag, prop, prop2)
    For count = 1 To up
        test(0, count) = Trim(test(0, count))
        test(1, count) = Trim(test(1, count))
        If InStr(test(1, count), "#") > 0 Then test(1, count) = Left(test(1, count), InStr(test(1, count), "#") - 1)
        If test(1, count) <> Empty Then
            temp2 = InStr(1, test(1, count), "mailto:", vbTextCompare)
            temp3 = 0
            If temp2 > 0 Then temp3 = InStr(temp2, test(1, count), "@", vbTextCompare)
            If temp2 = 0 And temp3 = 0 Then
                tempstr = chkurl(BASEHREF, test(1, count))
                additem lstsec, False, test(0, count), tempstr
                If useextention Then icon = GetExtention(tempstr)
                lstsec.ListItems.item(lstsec.ListItems.count).Checked = True
                lstsec.tag = GetIcon(icon, Me.imlico, Me.picmain)
                If imlico.ListImages.count = 1 Then Set lstsec.SmallIcons = imlico
                lstsec.ListItems.item(lstsec.ListItems.count).SmallIcon = Val(lstsec.tag)
            Else
                If chksec(4).value = vbChecked Then
                    additem lstsec, False, test(0, count), Right(test(1, count), Len(test(1, count)) - 7)
                    lstsec.tag = GetIcon(".eml", Me.imlico, Me.picmain)
                    If imlico.ListImages.count = 1 Then Set lstsec.SmallIcons = imlico
                    lstsec.ListItems.item(lstsec.ListItems.count).SmallIcon = Val(lstsec.tag)
                End If
            End If
            DoEvents
        End If
    Next
End Sub
Public Sub enumatag1(ByVal HTMLCode As String, BASEHREF As String, tag As String, prop As String, prop2 As String, icon As String)
    Dim test() As String, up As Long, higher As Long, count As Long
    higher = enumHTMLTAGS(HTMLCode, test, up, tag, prop2)
    For count = 1 To up
        test(0, count) = Trim(test(0, count))
        If test(0, count) <> Empty Then
            additem lstsec, False, prop, chkurl(BASEHREF, test(0, count))
            lstsec.ListItems.item(lstsec.ListItems.count).Checked = True
            lstsec.tag = GetIcon(test(0, count), Me.imlico, Me.picmain)
            If imlico.ListImages.count = 1 Then Set lstsec.SmallIcons = imlico
            lstsec.ListItems.item(lstsec.ListItems.count).SmallIcon = Val(lstsec.tag)
            DoEvents
        End If
    Next
End Sub

Public Sub enumimagelinks(url As String)
    Dim test() As String, up As Long, higher As Long, count As Long
    Dim HTMLCode As String, Filename As String, BASEHREF As String, IMGFILE As String, imlcount As Long
    Filename = chkfile(txtmain(1), "index.html")
    Filename = uniquefilename(Filename)
        
    Me.Caption = "Downloading HTML"
    If DownloadFile(url, Filename) Then
    HTMLCode = Replace(loadfile(Filename), vbNewLine, Empty)
    BASEHREF = getbasehref(url, HTMLCode)
    Me.Caption = "Scanning for Image Links"
    higher = enumHTMLTAGS(HTMLCode, test, up, "hrefimg", "href", "src")
    Me.Caption = "Enumerating Image Links"
    Me.Caption = up & " images found"
    
    If chksec(0).value = vbChecked Then enumatag1 HTMLCode, BASEHREF, "body", "background", "background", ".jpg"
    If chksec(1).value = vbChecked Then enumatag HTMLCode, BASEHREF, "img", "alt", "src", ".bmp", True
    If chksec(2).value = vbChecked Then enumatag HTMLCode, BASEHREF, "a", "node", "href", ".html", True
    If chksec(3).value = vbChecked Then enumatag HTMLCode, BASEHREF, "script", "language", "src", ".vbs", False
    
    removedoubles lstsec, 2
    autosizeall lstsec
    
    For count = 1 To up
    test(0, count) = Trim(test(0, count))
    test(1, count) = Trim(test(1, count))
    If test(0, count) <> Empty And test(1, count) <> Empty And cmdmain(8).Visible = True Then
        If isanimage(test(0, count)) Then
        additem lstmain, False, test(0, count), test(1, count), BASEHREF
        IMGFILE = chkfile(txtmain(1), makeabsolute(test(0, count)))
        IMGFILE = uniquefilename(IMGFILE)
        Call DownloadFile(chkurl(BASEHREF, test(1, count)), IMGFILE)
        'MsgBox chkurl(BASEHREF, test(1, count)), , IMGFILE
        imlcount = imlthumbs.ListImages.count
        
        If addpic(IMGFILE) Then
            If imlcount = 0 Then Set lstmain.SmallIcons = imlthumbs: Set lstmain.icons = imlthumbs
                If lstmain.ListItems.count > 0 Then
                With lstmain.ListItems.item(lstmain.ListItems.count)
                    .SmallIcon = imlcount + 1
                    .icon = imlcount + 1
                    .Checked = True
                End With
                End If
            End If
        End If
        If FileExists(IMGFILE) And IMGFILE <> Empty Then Kill IMGFILE
    
    Else
        'MsgBox test(0, count), , test(1, count)
    End If
    Me.Caption = "Downloaded " & count & " of " & up & " (" & Round(count / up * 100) & "%)"
    DoEvents
    Next
    'autosizeall lstmain
    Kill Filename
    End If
End Sub
Public Function addpic(IMGFILE As String) As Boolean
    On Error GoTo endit:
    Dim temp As Long
    With imlthumbs.ListImages
        temp = .count
        .Add , , LoadPicture(IMGFILE)
        If temp < .count Then addpic = True
    End With
endit:
End Function
Public Function getbasehref(url As String, ByVal HTMLCode As String) As String
    On Error Resume Next
    If InStrRev(url, "/") > InStr(url, "/") + 1 Then
        getbasehref = Left(url, InStrRev(url, "/") - 1)
    Else
        getbasehref = url
    End If
    
    If InStr(1, HTMLCode, "<base", vbTextCompare) > 0 Then
    Dim test() As String, up As Long, higher As Long, count As Long
    higher = enumHTMLTAGS(HTMLCode, test, up, "base", "href")
    For count = 1 To up
        If test(0, count) <> Empty Then getbasehref = test(0, count)
        DoEvents
    Next
    End If
End Function

Private Sub imgmain_Click()
    On Error Resume Next
    Dim temp As String, tempstr() As String, spoth As Long
    Dim url As String, lo As String, hi As String
    temp = Command
    If Len(temp) = 0 Then temp = InputBox("Please give me an EZCode", "EZCode entry")
    If Len(temp) = 0 Then Exit Sub
    tempstr = Split(temp, "?")
    url = Left(tempstr(0), InStrRev(tempstr(0), "/") - 1)
    lo = Right(tempstr(0), Len(tempstr(0)) - InStrRev(tempstr(0), "/"))
    hi = tempstr(1)
    
    If StrComp(killallceptnumber(lo, "|", True), killallceptnumber(hi, "|", True), vbTextCompare) = 0 Then
        txtmain(0) = url
        txtfiles = num2patt(lo)
        txtlbound = Val(killallceptnumber(lo))
        txtubound = Val(killallceptnumber(hi))
        'If Not isvalidpattern(txtfiles) Then MsgBox "This is not a valid pattern, all of the #'s must be put together", vbCritical, "Not a valid pattern"
    Else
        MsgBox "This is not a valid EZCode that I am capable of extracting", vbCritical, "Error in the code"
    End If
End Sub

Private Sub lstmain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then PopupMenu mnufile
End Sub

Private Sub lstsec_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstsec.Sorted = True
    lstsec.SortKey = ColumnHeader.Index
End Sub

Private Sub lstsec_DblClick()
    getitem lstsec.selecteditem.SmallIcon
End Sub

Public Sub getitem(Index As Long)
Select Case LCase(imlico.ListImages.item(Index).tag)
    Case ".htm", ".html", ".shtml", ".php", ".cgi", ".jsp"
        isclicking = True
        lstsec.selecteditem.ForeColor = vbBlue
        txtmain(0) = lstsec.selecteditem.SubItems(1)
        cmdmain_Click 4
        cmdmain_Click 0
        isclicking = False
End Select
End Sub
Private Sub lstsec_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then PopupMenu mnufile
End Sub

Public Sub selectsel(list As ListView, Optional auto As Boolean = True)
    Dim temp As Long
    For temp = 1 To list.ListItems.count
        If list.ListItems.item(temp).Selected Then list.ListItems.item(temp).Checked = auto
    Next
End Sub

Private Sub mnucheck_Click(Index As Integer)
If lstmain.Visible Then
    selectsel lstmain, Index = 0
Else
    selectsel lstsec, Index = 0
End If
End Sub

Private Sub txtmain_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Then
        Select Case KeyAscii
            Case 10: txtmain(0).text = "http://www." & txtmain(0).text & ".com" 'ctrl enter
            Case 13: cmdmain_Click 0 'enter
        End Select
    End If
End Sub
