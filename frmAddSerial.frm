VERSION 5.00
Begin VB.Form frmAddSerial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Serial Numbers"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmAddSerial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2910
   ScaleWidth      =   4695
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtProgram 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtNotes 
      Alignment       =   2  'Center
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1560
      Width           =   4455
   End
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   1200
      Width           =   3615
   End
   Begin VB.TextBox txtCode 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   3615
   End
   Begin VB.TextBox txtSerial 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Serial"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblCode 
      BackStyle       =   0  'Transparent
      Caption         =   "Code:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblSerialNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Serial #:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblProgram 
      BackStyle       =   0  'Transparent
      Caption         =   "Program: "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmAddSerial"
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
    Set rs = db.OpenRecordset("tblSerial")
    
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
        If txtProgram.Text = "" Then
            !Program = "None"
        Else
            !Program = txtProgram.Text
        End If
        If txtSerial.Text = "" Then
            !Serial = "None"
        Else
            !Serial = txtSerial.Text
        End If
        If txtCode.Text = "" Then
            !Code = "None"
        Else
            !Code = txtCode.Text
        End If
        If txtName.Text = "" Then
            !Name = "None"
        Else
            !Name = txtName.Text
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
    txtProgram.Text = ""
    txtSerial.Text = ""
    txtCode.Text = ""
    txtName.Text = ""
    txtNotes.Text = ""
End Sub

Private Sub cmdClose_Click()
    rs.Close
    db.Close
    Me.Hide
End Sub
