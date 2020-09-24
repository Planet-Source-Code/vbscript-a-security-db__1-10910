VERSION 5.00
Begin VB.Form frmAllPass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Passwords"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmAllPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   4695
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   2640
      Width           =   1335
   End
   Begin VB.ListBox lstRecords 
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View Record"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Close"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   2640
      Width           =   1455
   End
End
Attribute VB_Name = "frmAllPass"
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
    Set rs = db.OpenRecordset("tblPassword")
    
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
    cmdRefresh_Click
End Sub

Private Sub cmdView_Click()
    Dim intLoopIndex
    intLoopIndex = 0
    rs.MoveFirst

    For intLoopIndex = 0 To lstRecords.ListCount - 1
        If lstRecords.Selected(intLoopIndex) Then
            Do Until rs.EOF
                If rs.Fields("Server") Like lstRecords.Text Then
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

Private Sub cmdRefresh_Click()
    Dim Item As String
    lstRecords.Clear
    rs.MoveFirst
    Do While Not rs.EOF
        Item = rs.Fields("Server")
        lstRecords.AddItem Item
        rs.MoveNext
    Loop
End Sub

Private Sub lstRecords_DblClick()
    cmdView_Click
End Sub

Private Sub cmdExit_Click()
    Me.Hide
End Sub

Private Sub lstRecords_KeyPress(KeyAscii As Integer)
    cmdView_Click
End Sub
