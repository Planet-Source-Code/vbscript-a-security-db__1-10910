VERSION 5.00
Begin VB.Form frmChangePass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "   Change Password"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Change Password"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox txtNewPass2 
      Alignment       =   2  'Center
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtNewPass1 
      Alignment       =   2  'Center
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox txtOldPass 
      Alignment       =   2  'Center
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   3840
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   3840
      X2              =   120
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label lblConfirm 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblNewPass1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Password"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblOldPass 
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim db As Database
    Dim rs As Recordset
    
Private Sub Form_Load()
    Dim I As Integer, Y As Integer
    Dim lngVersion As Long
    lngVersion = GetVersion()
    If lngVersion = 143851525 Then
        SetLayered Me.hWnd, True, 175
    End If
    
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

Private Sub cmdOk_Click()
    Dim NewPass1, NewPass2, OldPass, Password
    
    Set db = OpenDatabase(DBPath)
    Set rs = db.OpenRecordset("tblLogin")
    rs.MoveFirst
    
    OldPass = txtOldPass.Text
    If OldPass = rs.Fields("Password") Then
        NewPass1 = txtNewPass1.Text
        NewPass2 = txtNewPass2.Text
        If NewPass1 = NewPass2 Then
            With rs
                .Edit
                !Password = NewPass1
                .Update
            End With
            MsgBox "Congradulations, your password has changed.", vbInformation + vbOKOnly, "Change Successfull"
        Else
            MsgBox "Passwords do not match, try again!", vbCritical + vbOKOnly, "Password Error"
        End If
    Else
        MsgBox "Passwords do not match, try again!", vbCritical + vbOKOnly, "Password Error"
    End If
    cmdCancel_Click
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Function Clear()
    txtOldPass.Text = ""
    txtNewPass1.Text = ""
    txtNewPass2.Text = ""
End Function
