VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "SecurityDB"
   ClientHeight    =   6030
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":08CA
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   5760
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7911
            Text            =   "Security Database"
            TextSave        =   "Security Database"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "8/21/2000"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "4:04 PM"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   7200
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuChangePass 
         Caption         =   "&Change Password"
      End
      Begin VB.Menu mnuOpenDatabase 
         Caption         =   "&Open Database"
      End
      Begin VB.Menu mnuspace00 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuData 
      Caption         =   "&Databases"
      Begin VB.Menu mnuSerial 
         Caption         =   "&Serial Numbers"
         Begin VB.Menu mnuViewSerial 
            Caption         =   "&View Serial Number"
         End
         Begin VB.Menu mnuSearchSerial 
            Caption         =   "&Search Serial Number"
         End
         Begin VB.Menu mnuspace 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAddSerial 
            Caption         =   "&Add Serial Number"
         End
         Begin VB.Menu mnuEditSerial 
            Caption         =   "&Edit Serial Number"
         End
         Begin VB.Menu mnuDelSerial 
            Caption         =   "&Delete Serial Number"
         End
      End
      Begin VB.Menu mnuPasswords 
         Caption         =   "&Passwords"
         Begin VB.Menu mnuViewPass 
            Caption         =   "&View Password"
         End
         Begin VB.Menu mnuSearchPasswords 
            Caption         =   "&Search Password"
         End
         Begin VB.Menu mnuSpace3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAddPass 
            Caption         =   "&Add Password"
         End
         Begin VB.Menu mnuEditPass 
            Caption         =   "&Edit Password"
         End
         Begin VB.Menu mnuDelPass 
            Caption         =   "&Delete Password"
         End
      End
      Begin VB.Menu mnuGetCache 
         Caption         =   "&Get Cached Passwords"
      End
   End
   Begin VB.Menu mnnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuContents 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuspace000 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Private Sub MDIForm_Load()
    Dim lngVersion As Long
    lngVersion = GetVersion()
    
    Me.Width = 7800
    Me.Height = 6720

    If (lngVersion = 143851525) Or ((lngVersion And &H80000000) = 0) Then
        mnuGetCache.Enabled = False
    Else
        mnuGetCache.Enabled = True
    End If

End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuChangePass_Click()
    frmChangePass.Show
End Sub

Private Sub mnuContents_Click()
    Dim nRun
    nRun = Shell("c:\winnt\hh.exe " & App.Path & "\SecurityDB.chm", vbMaximizedFocus)
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuOpenDatabase_Click()
    Dim sFile As String
    
    sFile = GetSetting("SecurityDB", "Database", "Path")
    
    With dlgCommonDialog
        .DialogTitle = "Open Database"
        .CancelError = False
        .Filter = "Databases (*.mdb)|*.mdb"
        .InitDir = sFile
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    SaveSetting "SecurityDB", "Database", "Path", sFile

End Sub

Private Sub mnuSearchPasswords_Click()
    frmSearchPass.Show
End Sub

Private Sub mnuSearchSerial_Click()
    frmSearchSerial.Show
End Sub

Private Sub mnuViewSerial_Click()
    frmAllSerial.Show
End Sub

Private Sub mnuAddSerial_Click()
    frmAddSerial.Show
End Sub

Private Sub mnuEditSerial_Click()
    frmEditSerial.Show
End Sub

Private Sub mnuDelSerial_Click()
    frmDelSerial.Show
End Sub

Private Sub mnuViewPass_Click()
    frmAllPass.Show
End Sub

Private Sub mnuAddPass_Click()
    frmAddPass.Show
End Sub

Private Sub mnuEditPass_Click()
    frmEditPass.Show
End Sub

Private Sub mnuDelPass_Click()
    frmDelPass.Show
End Sub

Private Sub mnuGetCache_Click()
    frmGetCachePass.Show
End Sub

Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub
