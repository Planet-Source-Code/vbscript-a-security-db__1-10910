VERSION 5.00
Begin VB.Form frmViewPass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Passwords"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmViewPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2550
   ScaleWidth      =   4695
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox txtRecord 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtNotes 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      Height          =   735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1200
      Width           =   4455
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   840
      Width           =   3495
   End
   Begin VB.TextBox txtUsername 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox txtServer 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblUsername 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblServer 
      BackStyle       =   0  'Transparent
      Caption         =   "Site/Server:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmViewPass"
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
    
    rs.MoveFirst
    GetData
End Sub

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Function GetData()
    txtRecord.Text = "Record " & rs.Fields("ID")
    txtServer.Text = rs.Fields("Server")
    txtUsername.Text = rs.Fields("Username")
    txtPassword.Text = rs.Fields("Password")
    txtNotes.Text = rs.Fields("Notes")
End Function
