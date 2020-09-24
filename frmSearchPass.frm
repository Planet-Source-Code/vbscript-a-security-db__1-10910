VERSION 5.00
Begin VB.Form frmSearchPass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search for Passwords"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optOr 
      Height          =   200
      Left            =   240
      TabIndex        =   12
      Top             =   2040
      Width           =   200
   End
   Begin VB.OptionButton optAnd 
      Height          =   200
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Value           =   -1  'True
      Width           =   200
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   3840
      Width           =   1335
   End
   Begin VB.ListBox lstResults 
      Height          =   1035
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   4455
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label lblOr 
      BackStyle       =   0  'Transparent
      Caption         =   "Use OR Logic"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblAnd 
      BackStyle       =   0  'Transparent
      Caption         =   "Use AND Logic"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Search on Username"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label lblSearch 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Search on Server/Site"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   4455
   End
End
Attribute VB_Name = "frmSearchPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim db As Database
    Dim rs As Recordset

Private Sub Form_Load()
    Dim I As Integer, Y As Integer
    
    Set db = OpenDatabase(DBPath)
    Set rs = db.OpenRecordset("tblPassword", dbOpenDynaset)
    Me.AutoRedraw = True
    Me.DrawStyle = 6
    Me.DrawMode = 13
    Me.DrawWidth = 13
    Me.ScaleMode = 3
    Me.ScaleHeight = 256
    
    For I = 0 To 510
        Me.Line (0, Y)-(Me.Width, Y + 1), RGB(0, 0, I), BF
        Y = Y + 1
    Next I
    ResetForm

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdReset_Click()
    ResetForm
End Sub

Private Sub cmdClear_Click()
    txtServer.Text = ""
    txtUser.Text = ""
    optAnd.Value = True
End Sub

Private Sub cmdSearch_Click()
    Dim strServer As String
    Dim strUser As String
    Dim strSearch As String
    
    lstResults.Clear
    
    strServer = "[Server] Like '*" & txtServer.Text & "*'"
    strUser = "[Username] Like '*" & txtUser.Text & "*'"
    
    If optAnd.Value = True And optOr.Value = False Then
        strSearch = strServer & " AND " & strUser
    ElseIf optAnd.Value = False And optOr.Value = True Then
        strSearch = strServer & " OR " & strUser
    End If
    
    With rs
        .FindFirst strSearch
        If .NoMatch Then
            MsgBox "Your search is complete.", vbInformation + vbOKOnly, "Search Complete"
        Else
            lstResults.AddItem (rs.Fields("Server"))
            Again
        End If
    End With
End Sub

Private Sub Again()
    Dim strSearch As String
    Dim strServer As String
    Dim strUser As String
    
    strServer = "[Server] Like '*" & txtServer.Text & "*'"
    strUser = "[Username] Like '*" & txtUser.Text & "*'"
    
    If optAnd.Value = True And optOr.Value = False Then
        strSearch = strServer & " AND " & strUser
    ElseIf optAnd.Value = False And optOr.Value = True Then
        strSearch = strServer & " OR " & strUser
    End If

    With rs
        .FindNext strSearch
        If .NoMatch Then
            MsgBox "Your search is complete.", vbInformation + vbOKOnly, "Search Complete"
        Else
            lstResults.AddItem (rs.Fields("Server"))
            Again
        End If
    End With
End Sub
Private Sub cmdView_Click()
    Dim intLoopIndex
    intLoopIndex = 0
    rs.MoveFirst

    For intLoopIndex = 0 To lstResults.ListCount - 1
        If lstResults.Selected(intLoopIndex) Then
            Do Until rs.EOF
                If rs.Fields("Server") Like lstResults.Text Then
                    frmViewPass.Show
                    frmViewPass.txtRecord.Text = "Record " & rs.Fields("ID")
                    frmViewPass.txtServer.Text = rs.Fields("Server")
                    frmViewPass.txtUsername.Text = rs.Fields("Username")
                    frmViewPass.txtPassword.Text = rs.Fields("Password")
                    frmViewPass.txtNotes.Text = rs.Fields("Notes")
                    Exit Sub
                Else
                    rs.MoveNext
                End If
            Loop
        End If
    Next intLoopIndex
End Sub

Private Sub lstResults_DblClick()
    cmdView_Click
End Sub

Function ResetForm()
    txtServer.Text = ""
    txtUser.Text = ""
    optAnd.Value = True
    lstResults.Clear
End Function
