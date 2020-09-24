VERSION 5.00
Begin VB.Form frmGetCachePass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get Passwords from Cache"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   Icon            =   "frmGetCache.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2190
   ScaleWidth      =   3870
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View Password"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.ListBox lstPasswords 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00000000&
      Height          =   1425
      ItemData        =   "frmGetCache.frx":0E42
      Left            =   120
      List            =   "frmGetCache.frx":0E44
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmGetCachePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    Dim I As Integer, y As Integer
    
    Me.AutoRedraw = True
    Me.DrawStyle = 6
    Me.DrawMode = 13
    Me.DrawWidth = 13
    Me.ScaleMode = 3
    Me.ScaleHeight = 256
    
    For I = 0 To 510
        Me.Line (0, y)-(Me.Width, y + 1), RGB(0, 0, I), BF
        y = y + 1
    Next I
    
    lstPasswords.Clear
    GetPasswds
End Sub
Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub cmdView_Click()
    Dim intLoopIndex

    For intLoopIndex = 0 To lstPasswords.ListCount - 1
        If lstPasswords.Selected(intLoopIndex) Then
            MsgBox lstPasswords.List(intLoopIndex), vbInformation + vbOKOnly, "Selected Password Information"
        End If
    Next intLoopIndex
End Sub

Private Sub lstPasswords_DblClick()
    cmdView_Click
End Sub
