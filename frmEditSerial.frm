VERSION 5.00
Begin VB.Form frmEditSerial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Serial Numbers"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmEditSerial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3270
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   960
      TabIndex        =   15
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtSerial 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   480
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
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   1200
      Width           =   3615
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
   Begin VB.CommandButton cmdStart 
      Caption         =   "|<<"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<<"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtRecord 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">>"
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   ">>|"
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   2760
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
   Begin VB.Label lblProgram 
      BackStyle       =   0  'Transparent
      Caption         =   "Program: "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblSerialNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Serial #:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblCode 
      BackStyle       =   0  'Transparent
      Caption         =   "Code:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   735
   End
End
Attribute VB_Name = "frmEditSerial"
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
    
    rs.MoveFirst
    GetData
End Sub

Private Sub cmdStart_Click()
    On Error Resume Next
    rs.MoveFirst
    GetData
End Sub

Private Sub cmdBack_Click()
    On Error Resume Next
    rs.MovePrevious
    GetData
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    rs.MoveNext
    GetData
End Sub

Private Sub cmdEnd_Click()
    On Error Resume Next
    rs.MoveLast
    GetData
End Sub

Private Sub cmdUpdate_Click()
    With rs
        .Edit
        !Program = txtProgram.Text
        !Serial = txtSerial.Text
        !Code = txtCode.Text
        !Name = txtName.Text
        !Notes = txtNotes.Text
        .Update
    End With
    MsgBox "The record was updated.", vbOKOnly + vbInformation, "Update Was Successfull"

End Sub

Private Sub cmdClose_Click()
    rs.Close
    db.Close
    Me.Hide
End Sub

Function GetData()
    txtRecord.Text = "Record " & rs.Fields("ID")
    txtProgram.Text = rs.Fields("Program")
    txtSerial.Text = rs.Fields("Serial")
    txtCode.Text = rs.Fields("Code")
    txtName.Text = rs.Fields("Name")
    txtNotes.Text = rs.Fields("Notes")
End Function
