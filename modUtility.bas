Attribute VB_Name = "modUtility"
Option Explicit

Public fMainForm As frmMain
Declare Function WNetEnumCachedPasswords Lib "mpr.dll" (ByVal s As String, ByVal I As Integer, ByVal b As Byte, ByVal proc As Long, ByVal l As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Declare Function UpdateLayeredWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
Public Const mlngWindows95 = 0
Public Const mlngWindowsNT = 1
Public Const mlngWindows2000 = 2
Public Declare Function GetVersion Lib "kernel32" () As Long
Public glngWhichWindows32 As Long

Type PASSWORD_CACHE_ENTRY
    cbEntry As Integer
    cbResource As Integer
    cbPassword As Integer
    iEntry As Byte
    nType As Byte
    abResource(1 To 1024) As Byte
    End Type

Public Type POINTAPI
    X As Long
    Y As Long
    End Type

Public Type SIZE
    cx As Long
    cy As Long
    End Type

Public Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
    End Type
    Public Const WS_EX_LAYERED = &H80000
    Public Const GWL_STYLE = (-16)
    Public Const GWL_EXSTYLE = (-20)
    Public Const AC_SRC_OVER = &H0
    Public Const AC_SRC_ALPHA = &H1
    Public Const AC_SRC_NO_PREMULT_ALPHA = &H1
    Public Const AC_SRC_NO_ALPHA = &H2
    Public Const AC_DST_NO_PREMULT_ALPHA = &H10
    Public Const AC_DST_NO_ALPHA = &H20
    Public Const LWA_COLORKEY = &H1
    Public Const LWA_ALPHA = &H2
    Public Const ULW_COLORKEY = &H1
    Public Const ULW_ALPHA = &H2
    Public Const ULW_OPAQUE = &H4
    Public lret As Long

Function CheckLayered(ByVal hWnd As Long) As Boolean
    lret = GetWindowLong(hWnd, GWL_EXSTYLE)
    If (lret And WS_EX_LAYERED) = WS_EX_LAYERED Then
        CheckLayered = True
    Else
        CheckLayered = False
    End If
End Function

Function SetLayered(ByVal hWnd As Long, SetAs As Boolean, bAlpha As Byte)
    lret = GetWindowLong(hWnd, GWL_EXSTYLE)
    If SetAs = True Then
        lret = lret Or WS_EX_LAYERED
    Else
        lret = lret And Not WS_EX_LAYERED
    End If
    SetWindowLong hWnd, GWL_EXSTYLE, lret
    SetLayeredWindowAttributes hWnd, 0, bAlpha, LWA_ALPHA
End Function

Public Function callback(X As PASSWORD_CACHE_ENTRY, ByVal lSomething As Long) As Integer
    Dim nLoop As Integer
    Dim cString As String
    Dim ccomputer
    Dim Resource As String
    Dim ResType As String
    Dim Password As String
    ResType = X.nType

    For nLoop = 1 To X.cbResource
        If X.abResource(nLoop) <> 0 Then
            cString = cString & Chr(X.abResource(nLoop))
        Else
            cString = cString & " "
        End If
    Next

    Resource = cString
    cString = ""

    For nLoop = X.cbResource + 1 To (X.cbResource + X.cbPassword)
        If X.abResource(nLoop) <> 0 Then
            cString = cString & Chr(X.abResource(nLoop))
        Else
            cString = cString & " "
        End If
    Next

    Password = cString
    cString = ""
    'for only the dialup passwords activate next line
    'If X.nType <> 6 Then GoTo 66
    frmGetCachePass.lstPasswords.AddItem " R: " & Resource & " P: " & Password
66
        callback = True
    End Function

Function GetPasswds()
    Dim nLoop As Integer
    Dim cString As String
    Dim lLong As Long
    Dim bByte As Byte
    bByte = &HFF
    nLoop = 0
    lLong = 0
    cString = ""
    Call WNetEnumCachedPasswords(cString, nLoop, bByte, AddressOf callback, lLong)
End Function
  
Sub Main()
    Dim fLogin As New frmLogin

    fLogin.Show vbModal
    If Not fLogin.OK Then
        End
    End If
    Unload fLogin
    Set fMainForm = New frmMain
    Load fMainForm
    fMainForm.Show

End Sub

Public Function DBPath() As String
    DBPath = GetSetting("SecurityDB", "Database", "Path")
    If DBPath = "" Then
        DBPath = InputBox("Please type the full path to your database.", "Database Not Found", App.Path & "\data.mdb")
        SaveSetting "SecurityDB", "Database", "Path", DBPath
    End If
End Function
