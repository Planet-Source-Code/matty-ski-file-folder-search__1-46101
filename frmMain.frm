VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Search..."
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14055
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   14055
   Begin VB.Timer tmrCount 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3000
      Top             =   1920
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   615
      Left            =   3000
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   16
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   3000
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView View 
      Height          =   4815
      Left            =   2880
      TabIndex        =   15
      Top             =   120
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   8493
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "In Folder"
         Object.Width           =   5733
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Size"
         Object.Width           =   1976
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date Modified"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Type"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Attributes"
         Object.Width           =   1482
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "S&earch Parameters"
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.CheckBox chkNoShow 
         Caption         =   "Do not show icons"
         Height          =   375
         Left            =   1320
         TabIndex        =   14
         ToolTipText     =   "It loads the files a lot faster into the Viewer to the right"
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "&Start"
         Default         =   -1  'True
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Ê"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   4
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "É"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   3
         Top             =   960
         Width           =   375
      End
      Begin VB.ComboBox cmbLocation 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   4200
         Width           =   2415
      End
      Begin VB.ComboBox cmbPhrase 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   3480
         Width           =   2415
      End
      Begin VB.ListBox lstFiles 
         Height          =   1425
         Index           =   0
         ItemData        =   "frmMain.frx":0D1C
         Left            =   120
         List            =   "frmMain.frx":0D1E
         TabIndex        =   7
         ToolTipText     =   "Can create duplicate files"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.ListBox lstFiles 
         Height          =   1425
         Index           =   1
         ItemData        =   "frmMain.frx":0D20
         Left            =   1320
         List            =   "frmMain.frx":0D22
         TabIndex        =   8
         Top             =   1680
         Width           =   1095
      End
      Begin VB.ComboBox cmbFileName 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label lbl 
         Caption         =   "Files to Exclude:"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   6
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lbl 
         Caption         =   "Files to Include:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lbl 
         Caption         =   "Start Location:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   11
         Top             =   3960
         Width           =   2295
      End
      Begin VB.Label lbl 
         Caption         =   "Phrase contained in the files:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Label lbl 
         Caption         =   "All or Part of the file name:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Label lblStat 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4680
      TabIndex        =   18
      Top             =   5040
      Width           =   9255
   End
   Begin VB.Label lblCount 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0 Files Found"
      Height          =   255
      Left            =   2865
      TabIndex        =   17
      Top             =   5040
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 260
    szTypeName As String * 80
End Type
Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(128) As Byte
End Type
Private Type OVERLAPPED
    Internal As Long
    InternalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type
Private Declare Function GetWindowsDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Function FindFirstFileA Lib "kernel32" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFileA Lib "kernel32" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
'Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
'Private Declare Function GetSystemDefaultLangID Lib "kernel32" () As Integer
'Private Declare Function GetSystemTimeAdjustment Lib "kernel32" (lpTimeAdjustment As Long, lpTimeIncrement As Long, lpTimeAdjustmentDisabled As Long) As Long
Private Declare Function SHGetFileInfoA Lib "Shell32" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
'Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Boolean
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal I&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal Flags&) As Long ' Copied from ?
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
'Private Declare Function ReadFileEx Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpOverlapped As OVERLAPPED, ByVal lpCompletionRoutine As Long) As Long
'Private Declare Function GetFileType Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Dim WindowsDir As String
Dim HoursOffGMT As Long      ' System time off from GMT time
Dim TotalDirList() As String
Dim CurrentDir As Long
Dim IconCount As Long
Dim FilesFound As Long
Dim StopSearch As Boolean

Private Sub cmbFileName_KeyPress(KeyAscii As Integer)
    Select Case Chr$(KeyAscii)
    Case "\", "/", Chr$(34), "(", ")", ":", "<", ">": Beep: KeyAscii = 0
    End Select
End Sub

Private Sub cmdAdd_Click(Index As Integer)
    Dim I As Long
    Dim IsInList As Boolean
    
    cmbFileName.Text = Trim$(cmbFileName.Text)
    If cmbFileName.Text = "" Then Exit Sub
    
    ' Wildcard check for Index 1
    If Index = 1 Then
        If InStr(1, cmbFileName.Text, "*") Or InStr(1, cmbFileName.Text, "?") Then
            MsgBox "Sorry, no wildcards for 'Files to Exclude'" & vbCr & "Only All / Part of the file name", vbExclamation
            Exit Sub
        End If
    End If
    
    ' Add entry to dropdown list
    For I = 0 To cmbFileName.ListCount - 1
        If UCase$(cmbFileName.List(I)) = UCase$(cmbFileName.Text) Then IsInList = True: Exit For
    Next I
    If IsInList = False Then cmbFileName.AddItem cmbFileName.Text Else IsInList = False
    
    ' Add entry & check for duplicates to long list
    For I = 0 To lstFiles(Index).ListCount - 1
        If UCase$(lstFiles(Index).List(I)) = UCase$(cmbFileName.Text) Then IsInList = True: Exit For
    Next I
    If IsInList = False Then lstFiles(Index).AddItem cmbFileName.Text
    
    cmdAdd(Index Xor 1).Enabled = False
    cmbFileName.SelStart = 0
    cmbFileName.SelLength = Len(cmbFileName.Text)
    cmbFileName.SetFocus
End Sub

Private Sub cmdStart_Click()
    On Error Resume Next ' For a file name used as a start location
    Dim I As Long
    Dim IsInList As Boolean
    
    ' Check to see if the search has already started
    If cmdStart.Caption = "&Stop" Then
        StopSearch = True
        Exit Sub
    End If
    
    ' Validates the 'Start Location' and check for a File Name, Clearups and Cosmetics
    ReDim TotalDirList(0)
    CurrentDir = 0
    IconCount = 0
    FilesFound = 0
    tmrCount.Enabled = True
    StopSearch = False
    cmbFileName.Text = Trim$(cmbFileName.Text)
    cmbPhrase.Text = Trim$(cmbPhrase.Text)
    cmbLocation.Text = Trim$(cmbLocation.Text)
    If Right$(cmbLocation.Text, 1) = "\" Then TotalDirList(0) = cmbLocation.Text Else TotalDirList(0) = cmbLocation.Text & "\"
    If cmbLocation.Text = "" Or Dir(TotalDirList(0), vbDirectory + vbHidden + vbSystem) = "" Then MsgBox "Invalid Start Location", vbExclamation, "Search": Exit Sub ' On Error
    View.Sorted = False
    View.SortOrder = 1 ' Sets Decending for the header click
    cmbFileName.Enabled = False
    cmbPhrase.Enabled = False
    cmbLocation.Enabled = False
    For I = 0 To 1
        If cmdAdd(I).Enabled = False Then cmdAdd(I).Tag = "X" Else cmdAdd(I).Enabled = False: cmdAdd(I).Tag = ""
    Next I
    cmdStart.Caption = "&Stop"
    
    ' Add typed entries to the lists
    If cmbFileName.Text <> "" Then
        For I = 0 To cmbFileName.ListCount - 1
            If UCase$(cmbFileName.List(I)) = UCase$(cmbFileName.Text) Then IsInList = True: Exit For
        Next I
        If IsInList = False Then cmbFileName.AddItem cmbFileName.Text Else IsInList = False
    End If
    If cmbPhrase.Text <> "" Then
        For I = 0 To cmbPhrase.ListCount - 1
            If UCase$(cmbPhrase.List(I)) = UCase$(cmbPhrase.Text) Then IsInList = True: Exit For
        Next I
        If IsInList = False Then cmbPhrase.AddItem cmbPhrase.Text Else IsInList = False
    End If
    For I = 0 To cmbLocation.ListCount - 1
        If UCase$(cmbLocation.List(I)) = UCase$(cmbLocation.Text) Then IsInList = True: Exit For
    Next I
    If IsInList = False Then cmbLocation.AddItem cmbLocation.Text
    
    ' Clean up the 'View' and the 'ImageList' icons
    View.ListItems.Clear
    View.SmallIcons = Nothing
    For I = 2 To ImageList.ListImages.Count
        ImageList.ListImages.Remove 2 ' Keeps the first folder icon
    Next I
    View.SmallIcons = ImageList
    
    Do
        ' Get a Directory list for searching the Sub Directories
        Dim hndDir As Long
        Dim FndDirDat As WIN32_FIND_DATA
        Dim FileName As String, LastFile As String
        hndDir = FindFirstFileA(TotalDirList(CurrentDir) & "*.*", FndDirDat) ' For "."
        If Left$(FndDirDat.cFileName, 2) = "." & Chr$(0) Then FindNextFileA hndDir, FndDirDat ' For ".."
        If Left$(FndDirDat.cFileName, 3) = ".." & Chr$(0) Then FindNextFileA hndDir, FndDirDat
        FileName = Mid$(FndDirDat.cFileName, 1, InStr(1, FndDirDat.cFileName, Chr$(0)) - 1)
        Do
            If (16 Xor FndDirDat.dwFileAttributes) < FndDirDat.dwFileAttributes And FileName <> ".." Then
                ReDim Preserve TotalDirList(UBound(TotalDirList()) + 1)
                TotalDirList(UBound(TotalDirList())) = TotalDirList(CurrentDir) & FileName & "\"
            End If
            LastFile = FileName
            FindNextFileA hndDir, FndDirDat
            FileName = Mid$(FndDirDat.cFileName, 1, InStr(1, FndDirDat.cFileName, Chr$(0)) - 1)
        Loop Until FileName = LastFile Or StopSearch = True
        FindClose hndDir
        
        ' Do the Search for the Files in a Listbox and Dir 'TotalDirList(CurrentDir)'
        lblStat.Caption = " Searching: " & TotalDirList(CurrentDir)
        If lstFiles(0).ListCount > 0 Then
            For I = 0 To lstFiles(0).ListCount - 1
                DoTheDir "*" & lstFiles(0).List(I) & "*"
            Next I
        ElseIf lstFiles(1).ListCount > 0 Then
            DoTheDir "*.*"
        Else
            If cmbFileName.Text = "" Then
                DoTheDir "*.*"
            Else
                DoTheDir cmbFileName.Text
            End If
        End If
        
        CurrentDir = CurrentDir + 1
    Loop Until UBound(TotalDirList()) + 1 = CurrentDir Or StopSearch = True
    
    ' Clean up Cosmetics
    cmbFileName.Enabled = True
    cmbPhrase.Enabled = True
    cmbLocation.Enabled = True
    tmrCount.Enabled = False
    tmrCount_Timer
    lblStat.Caption = ""
    For I = 0 To 1
        If cmdAdd(I).Tag = "" Then cmdAdd(I).Enabled = True
    Next I
    cmdStart.Caption = "&Start"
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim I As Long
    Dim FileNameVal As String, PhraseVal As String, LocationVal As String
    
    ' Loads in all the previous search bits that were typied
    Do
        I = I + 1
        FileNameVal = GetSetting(App.Title, "Lists", "File Name " & Format(I, "0000"), "")
        If FileNameVal <> "" Then cmbFileName.AddItem FileNameVal
        PhraseVal = GetSetting(App.Title, "Lists", "Phrase " & Format(I, "0000"), "")
        If PhraseVal <> "" Then cmbPhrase.AddItem PhraseVal
        LocationVal = GetSetting(App.Title, "Lists", "Location " & Format(I, "0000"), "")
        If LocationVal <> "" Then cmbLocation.AddItem LocationVal
    Loop Until FileNameVal = "" And PhraseVal = "" And LocationVal = ""
    
    chkNoShow.Value = GetSetting(App.Title, "Settings", "Dont Show Icons", 0)
    Me.Left = GetSetting(App.Title, "Settings", "Main Left", 0)
    Me.Top = GetSetting(App.Title, "Settings", "Main Top", 0)
    
    ' Check to see if quotations are in the parameters
    If Left(Command$, 1) = Chr$(34) Then
        cmbLocation = Mid$(Command$, 2, Len(Command$) - 2)
    Else
        cmbLocation = Command$
    End If
    
    ' Set Windows Dir
    Dim tmpDir As String * 260
    GetWindowsDirectoryA tmpDir, 260
    WindowsDir = Left$(tmpDir, InStr(1, tmpDir, Chr$(0)) - 1) & "\"
    
    ' Get hour difference from GMT
    Dim SysTime As SYSTEMTIME, Hour As Integer
    GetSystemTime SysTime
    HoursOffGMT = CInt(Format(Time, "HH")) - SysTime.wHour
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    ' Like the sound of heads banging against wood
    View.Width = Me.Width - 3120
    View.Height = Me.Height - 1080
    lblCount.Top = Me.Height - 900
    lblStat.Top = Me.Height - 900
    lblStat.Width = Me.Width - 4890
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim I As Long
    
    Do
        I = I + 1
        If I <= cmbFileName.ListCount Then SaveSetting App.Title, "Lists", "File Name " & Format(I, "0000"), cmbFileName.List(I - 1)
        If I <= cmbPhrase.ListCount Then SaveSetting App.Title, "Lists", "Phrase " & Format(I, "0000"), cmbPhrase.List(I - 1)
        If I <= cmbLocation.ListCount Then SaveSetting App.Title, "Lists", "Location " & Format(I, "0000"), cmbLocation.List(I - 1)
    Loop Until I > cmbFileName.ListCount And I > cmbPhrase.ListCount And I > cmbLocation.ListCount
    
    SaveSetting App.Title, "Settings", "Dont Show Icons", chkNoShow.Value
    If Me.WindowState = 0 Then
        SaveSetting App.Title, "Settings", "Main Left", Me.Left
        SaveSetting App.Title, "Settings", "Main Top", Me.Top
    End If
    StopSearch = True
    DoEvents ' This might help a little to close the files before terminating ?
    End
End Sub

Private Function DoTheDir(FileType As String)
    On Error Resume Next ' In case
    Dim hndFile As Long, I As Long, hndPic As Long
    Dim FndFleDat As WIN32_FIND_DATA
    Dim FileName As String, LastFile As String
    Dim SkipFile As Boolean
    Dim TheRealTime As SYSTEMTIME
    Dim SHFI As SHFILEINFO
    Dim ViewText As String
    
    ' Lists the Directory Contents according to 'FileType'
    hndFile = FindFirstFileA(TotalDirList(CurrentDir) & FileType, FndFleDat) ' For "."
    If Left$(FndFleDat.cFileName, 2) = "." & Chr$(0) Then FindNextFileA hndFile, FndFleDat ' For ".."
    If Left$(FndFleDat.cFileName, 3) = ".." & Chr$(0) Then FindNextFileA hndFile, FndFleDat
    FileName = Mid$(FndFleDat.cFileName, 1, InStr(1, FndFleDat.cFileName, Chr$(0)) - 1)
    Do
        If lstFiles(1).ListCount > 0 Then
            SkipFile = False
            For I = 0 To lstFiles(1).ListCount - 1
                If InStr(1, FileName, lstFiles(1).List(I), vbTextCompare) > 0 Then SkipFile = True: Exit For
            Next I
        End If
        If FileName <> ".." And FileName <> "" And SkipFile = False Then
            If SearchFileContents(TotalDirList(CurrentDir) & FileName) = True Then
                IconCount = IconCount + 1
                
                ' Check if the folder does not have an icon assosiated with it - and add a new Item
                If (16 Xor FndFleDat.dwFileAttributes) < FndFleDat.dwFileAttributes And (1 Xor FndFleDat.dwFileAttributes) > FndFleDat.dwFileAttributes Then
                    If chkNoShow.Value = 0 Then
                        View.ListItems.Add IconCount, , FileName, , 1
                    Else
                        View.ListItems.Add IconCount, , FileName
                    End If
                Else
                    If chkNoShow.Value = 0 Then
                        ' This way is wastefull on the memory
                        hndPic = SHGetFileInfoA(TotalDirList(CurrentDir) & FileName, 0&, SHFI, Len(SHFI), 26116)
                        picIcon.Cls
                        ImageList_Draw hndPic, SHFI.iIcon, picIcon.hDC, 0, 0, 1
                        ImageList.ListImages.Add , , picIcon.Image
                        View.ListItems.Add IconCount, , FileName, , ImageList.ListImages.Count
                    Else
                        View.ListItems.Add IconCount, , FileName
                    End If
                End If
                
                FilesFound = FilesFound + 1
                
                ' Add viewable data to the 'View'
                View.ListItems(IconCount).SubItems(1) = TotalDirList(CurrentDir)
                View.ListItems(IconCount).SubItems(2) = Format(Fix(CDbl("&H" & Hex(FndFleDat.nFileSizeHigh) & Hex(FndFleDat.nFileSizeLow)) / 1024), "#,##0") & " KB"
                FileTimeToSystemTime FndFleDat.ftLastWriteTime, TheRealTime
                View.ListItems(IconCount).SubItems(3) = Format(TheRealTime.wDay, "00") & "/" & Format(TheRealTime.wMonth, "00") & "/" & TheRealTime.wYear & " " & Format(TheRealTime.wHour + HoursOffGMT, "00") & ":" & Format(TheRealTime.wMinute, "00") & ":" & Format(TheRealTime.wSecond, "00") ' Can be inaccurate
                If (16 Xor FndFleDat.dwFileAttributes) < FndFleDat.dwFileAttributes Then
                    View.ListItems(IconCount).SubItems(4) = "Folder"
                Else
                    '    ' Get the file extention
                    '    For I = Len(View.SelectedItem.Text) To 1 Step -1
                    '        If InStr(I, View.SelectedItem.Text, ".") Then
                    '            Extention = UCase$(Mid$(View.SelectedItem.Text, I + 1))
                    '            Exit For
                    '        ElseIf InStr(I, View.SelectedItem.Text, "\") Then
                    '            Exit For
                    '        End If
                    '    Next I
                    View.ListItems(IconCount).SubItems(4) = "File"
                End If
                View.ListItems(IconCount).SubItems(5) = GetAttribs(FndFleDat.dwFileAttributes)
                'a = FndFleDat.cAlternate
                'b = FndFleDat.dwReserved0
                'c = FndFleDat.dwReserved1
                
            End If
        End If
        DoEvents
        
        LastFile = FileName
        FindNextFileA hndFile, FndFleDat
        FileName = Mid$(FndFleDat.cFileName, 1, InStr(1, FndFleDat.cFileName, Chr$(0)) - 1)
        If IconCount = 0 Then ViewText = "" Else ViewText = View.ListItems(IconCount).Text
    Loop Until ViewText = FileName Or FileName = ".." Or LastFile = FileName Or StopSearch = True ' Need to trim
    FindClose hndFile
End Function

Private Function GetAttribs(AttVal As Long) As String
    If (32 Xor AttVal) < AttVal Then GetAttribs = "A"
    'If (16 Xor AttVal) < AttVal Then GetAttribs = GetAttribs & "D" ' Directory
    If (8 Xor AttVal) < AttVal Then GetAttribs = GetAttribs & "R"
    If (4 Xor AttVal) < AttVal Then GetAttribs = GetAttribs & "S"
    If (2 Xor AttVal) < AttVal Then GetAttribs = GetAttribs & "H"
    If (1 Xor AttVal) < AttVal Then GetAttribs = GetAttribs & "I" ' Iconed
End Function

Private Function SearchFileContents(TheFileName As String) As Boolean
    On Error Resume Next ' In case
    If cmbPhrase.Text = "" Then SearchFileContents = True: Exit Function
    
    Dim LenPhrase As Long
    Dim OFS As OFSTRUCT
    Dim hndFile As Long
    Dim OL As OVERLAPPED
    Dim NOBR As Long
    Dim tmpBuff As String * 32768
    Dim ReadString As String
    Dim CR As Long
    
    ' Theres a 2GB read limit - thats enough
    LenPhrase = Len(cmbPhrase.Text) - 1
    hndFile = OpenFile(TheFileName, OFS, 0)
    Do
        ReadString = tmpBuff
        ReadFile hndFile, ByVal tmpBuff, 32768, NOBR, OL
        ReadString = Right$(ReadString, LenPhrase) & tmpBuff
        If InStr(1, ReadString, cmbPhrase.Text, vbTextCompare) > 0 Then SearchFileContents = True: Exit Do
        OL.offset = OL.offset + 32768
        DoEvents
    Loop Until NOBR < 32768 Or OL.offset > 2147450878 Or StopSearch = True
    CloseHandle hndFile
End Function

Private Sub lstFiles_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 46 Then lstFiles(Index).RemoveItem lstFiles(Index).ListIndex
    If lstFiles(Index).ListCount = 0 Then cmdAdd(Index Xor 1).Enabled = True
End Sub

Private Sub tmrCount_Timer()
    lblCount.Caption = Format(FilesFound, "#,###,##0") & " Files Found"
End Sub

Private Sub View_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ' Flawed with Size and Date
    If cmdStart.Caption = "&Stop" Then Exit Sub
    
    If View.SortKey = ColumnHeader.Index - 1 Then View.SortOrder = View.SortOrder Xor 1 Else View.SortOrder = 0
    View.SortKey = ColumnHeader.Index - 1
    View.Sorted = True
End Sub

Private Sub View_DblClick()
    On Error Resume Next
    Dim I As Long
    Dim Extention As String, MainProg As String, ProgFiles As String
    
    ' Check to see if its a folder and open it
    If View.ListItems(View.SelectedItem.Index).SubItems(4) = "Folder" Then
        Shell WindowsDir & "Explorer.exe /e, " & View.ListItems(View.SelectedItem.Index).SubItems(1) & View.SelectedItem.Text, vbNormalFocus
        Exit Sub
    End If
    
    ' Get the file extention
    For I = Len(View.SelectedItem.Text) To 1 Step -1
        If InStr(I, View.SelectedItem.Text, ".") Then
            Extention = UCase$(Mid$(View.SelectedItem.Text, I + 1))
            Exit For
        ElseIf InStr(I, View.SelectedItem.Text, "\") Then
            Exit For
        End If
    Next I
    
    ProgFiles = Left$(WindowsDir, 3) & "Program Files\"
    
    ' Load the file - These need to be adjusted, or a compleatly new routine is needed - something better than reading the registry
    Select Case Extention
    Case "EXE", "BAT", "COM", "CMD"
    'Case "LNK": MainProg = "RunDll32.exe "
    Case "VBP", "VBG", "FRM", "BAS", "CLS": MainProg = ProgFiles & "Microsoft Visual Studio\VB98\VB6.EXE " ' Remember to add a space at the end
    Case "DSP", "DSW", "C", "CPP", "H", "RC", "IDL", "TLB", "OBJ", "RES": MainProg = ProgFiles & "Microsoft Visual Studio\Common\MSDev98\Bin\MSDEV.EXE " ' "MAK" ?
    Case "DOC": MainProg = ProgFiles & "Microsoft Office\Office10\WINWORD.EXE "
    Case "XLS": MainProg = ProgFiles & "Microsoft Office\Office10\EXCEL.EXE "
    Case "PPT": MainProg = ProgFiles & "Microsoft Office\Office10\POWERPNT.EXE "
    Case "MDB": MainProg = ProgFiles & "Microsoft Office\Office10\MSACCESS.EXE "
    Case "HTM", "HTML", "ASP": MainProg = ProgFiles & "Internet Explorer\IExplore.exe "
    Case "TXT", "INI", "INF", "LOG", "REG", "VBW", "DEF", "JS", "VBS": MainProg = WindowsDir & "Notepad.exe "
    Case "MP3", "WAV": MainProg = ProgFiles & "Winamp\Winamp.exe "
    Case "WMV", "WMA", "ASF": MainProg = ProgFiles & "Windows Media Player\wmplayer.exe "
    Case "ZIP", "CAB", "Z", "GZ", "TAR": MainProg = ProgFiles & "Winzip\WINZIP32.EXE "
    Case "RAR", "ACE", "ISO": MainProg = ProgFiles & "WinRAR\WinRAR.exe "
    Case "JPG", "JPEG", "GIF", "BMP": MainProg = WindowsDir & "System32\MSPaint.exe " ' Yeah Win 2K+
    Case "HLP": MainProg = WindowsDir & "winhlp32.exe "
    Case "PDF": MainProg = ProgFiles & "Acrobat 5.0\Reader\AcroRd32.exe "
    Case Else: Exit Sub
    End Select
    ChDir View.ListItems(View.SelectedItem.Index).SubItems(1)
    Shell MainProg & Chr$(34) & View.ListItems(View.SelectedItem.Index).SubItems(1) & View.SelectedItem.Text & Chr$(34), vbNormalFocus
    If Err Then MsgBox "Error Loading Program," & vbCr & "Check the Sub 'View_DblClick'", vbExclamation
End Sub

Private Sub View_KeyUp(KeyCode As Integer, Shift As Integer)
    ' Delete the file if its not searching
    If KeyCode = 46 And cmdStart.Caption = "&Start" And View.ListItems.Count > 0 Then
        Dim DelFile As String
        DelFile = View.SelectedItem.SubItems(1) & View.SelectedItem.Text
        If MsgBox("Are you sure you want to delete" & vbCr & DelFile, vbYesNo + vbDefaultButton2 + vbExclamation, "Delete") = vbYes Then
            Kill DelFile
            View.ListItems.Remove View.SelectedItem.Index
        End If
    End If
End Sub
