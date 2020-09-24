VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About MyApp"
   ClientHeight    =   4080
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrintAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2816.089
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   3105
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   4260
      TabIndex        =   1
      Top             =   3555
      Width           =   1245
   End
   Begin VB.Image Image2 
      Height          =   600
      Left            =   600
      Picture         =   "frmPrintAbout.frx":000C
      Stretch         =   -1  'True
      Top             =   600
      Width           =   600
   End
   Begin VB.Line Line1 
      X1              =   112.686
      X2              =   5296.251
      Y1              =   1739.349
      Y2              =   1739.349
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   120
      Picture         =   "frmPrintAbout.frx":08D6
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "App Description"
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   1800
      TabIndex        =   2
      Top             =   1320
      Width           =   3735
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1800
      TabIndex        =   4
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label lblVersion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Version"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label lblDisclaimer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Warning: ..."
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   3855
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
        KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
        KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1    ' Unicode nul terminated string
Const REG_DWORD = 4    ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
    Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    lblDescription.Caption = App.Title & " from" + vbCrLf + vbCrLf + _
            ""
    lblDisclaimer.Caption = "Warning: This computer program is protected by copyright law and international treaties." + _
            "Unauthorized reproduction or distribution of this program, or any portion of it, may result in severe " + _
            "civil and criminal penalties, and will be prosecuted to the maximum extent possible under law."
    lblVersion.BorderStyle = 0
    lblTitle.BorderStyle = 0
    lblDescription.BorderStyle = 0
    lblDisclaimer.BorderStyle = 0
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr

    Dim rc          As Long
    Dim SysInfoPath As String

    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"

            ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
        ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If

    Call Shell(SysInfoPath, vbNormalFocus)

    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i           As Long    ' Loop Counter
    Dim rc          As Long    ' Return Code
    Dim hKey        As Long    ' Handle To An Open Registry Key
    Dim hDepth      As Long    '
    Dim KeyValType  As Long    ' Data Type Of A Registry Key
    Dim tmpVal      As String    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize  As Long    ' Size Of Registry Key Variable

    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)    ' Open Registry Key

    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError    ' Handle Error...

    tmpVal = String$(1024, 0)    ' Allocate Variable Space
    KeyValSize = 1024    ' Mark Variable Size

    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
            KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value

    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError    ' Handle Errors

    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then    ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)    ' Null Found, Extract From String
    Else    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)    ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType    ' Search Data Types...
        Case REG_SZ    ' String Registry Key Data Type
            KeyVal = tmpVal    ' Copy String Value
        Case REG_DWORD    ' Double Word Registry Key Data Type
            For i = Len(tmpVal) To 1 Step -1    ' Convert Each Bit
                KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))    ' Build Value Char. By Char.
            Next
            KeyVal = Format$("&h" + KeyVal)    ' Convert Double Word To String
    End Select

    GetKeyValue = True    ' Return Success
    rc = RegCloseKey(hKey)    ' Close Registry Key
    Exit Function    ' Exit

GetKeyError:        ' Cleanup After An Error Has Occured...
    KeyVal = ""    ' Set Return Val To Empty String
    GetKeyValue = False    ' Return Failure
    rc = RegCloseKey(hKey)    ' Close Registry Key
End Function
