VERSION 5.00
Begin VB.Form frmAddPass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Passwords"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmAddPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2535
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtServer 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.TextBox txtUsername 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   840
      Width           =   3495
   End
   Begin VB.TextBox txtNotes 
      Alignment       =   2  'Center
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1200
      Width           =   4455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblServer 
      BackStyle       =   0  'Transparent
      Caption         =   "Site/Server:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblUsername 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "frmAddPass"
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
End Sub

Private Sub cmdAdd_Click()
    With rs
        .AddNew
        If txtServer.Text = "" Then
            !Server = "None"
        Else
            !Server = txtServer.Text
        End If
        If txtUsername.Text = "" Then
            !UserName = "None"
        Else
            !UserName = txtUsername.Text
        End If
        If txtPassword.Text = "" Then
            !Password = "None"
        Else
            !Password = txtPassword.Text
        End If
        If txtNotes.Text = "" Then
            !Notes = "None"
        Else
            !Notes = txtNotes.Text
        End If
        .Update
    End With
    MsgBox "The record was added.", vbOKOnly + vbInformation, "Add Was Successfull"
    cmdClear_Click

End Sub

Private Sub cmdClear_Click()
    txtServer.Text = ""
    txtUsername.Text = ""
    txtPassword.Text = ""
    txtNotes.Text = ""
End Sub

Private Sub cmdClose_Click()
    rs.Close
    db.Close
    Me.Hide
End Sub


