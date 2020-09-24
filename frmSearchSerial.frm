VERSION 5.00
Begin VB.Form frmSearchSerial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search for Serial Numbers"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.ListBox lstResults 
      Height          =   2010
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   4455
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label lblSearch 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Search on Program Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   4455
   End
End
Attribute VB_Name = "frmSearchSerial"
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
    Set rs = db.OpenRecordset("tblSerial", dbOpenDynaset)
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
End Sub

Private Sub cmdSearch_Click()
    Dim strSearch As String
    
    lstResults.Clear
    
    strSearch = "[Program] Like '*" & txtSearch.Text & "*'"
    
    With rs
        .FindFirst strSearch
        If .NoMatch Then
            MsgBox txtSearch.Text & " not found.  Please try again.", vbInformation + vbOKOnly, "Record Not Found"
        Else
            lstResults.AddItem (rs.Fields("Program"))
            Again
        End If
    End With
End Sub

Sub Again()
    Dim strSearch As String
    
    strSearch = "[Program] Like '*" & txtSearch.Text & "*'"

    With rs
        .FindNext strSearch
        If .NoMatch Then
            MsgBox "Your search is complete.", vbInformation + vbOKOnly, "Search Complete"
        Else
            lstResults.AddItem (rs.Fields("Program"))
            Again
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    txtSearch.Text = ""
End Sub

Private Sub cmdReset_Click()
    lstResults.Clear
    txtSearch.Text = ""
End Sub
    
Private Sub cmdView_Click()
    Dim intLoopIndex
    intLoopIndex = 0
    rs.MoveFirst

    For intLoopIndex = 0 To lstResults.ListCount - 1
        If lstResults.Selected(intLoopIndex) Then
            Do Until rs.EOF
                If rs.Fields("Program") Like lstResults.Text Then
                    frmViewSerial.Show
                    frmViewSerial.txtRecord.Text = "Record " & rs.Fields("ID")
                    frmViewSerial.txtProgram.Text = rs.Fields("Program")
                    frmViewSerial.txtSerial.Text = rs.Fields("Serial")
                    frmViewSerial.txtCode.Text = rs.Fields("Code")
                    frmViewSerial.txtName.Text = rs.Fields("Name")
                    frmViewSerial.txtNotes.Text = rs.Fields("Notes")
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

