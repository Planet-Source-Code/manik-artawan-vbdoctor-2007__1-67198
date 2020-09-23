VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form SystemControl 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONTROL PANELS"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10380
   Icon            =   "SYSCONTROL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   10380
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1800
      Top             =   1800
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\VBDOCTOR\DataSystem.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\VBDOCTOR\DataSystem.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from label"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "About It"
      Height          =   855
      Left            =   120
      TabIndex        =   27
      Top             =   3960
      Width           =   4815
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright (C) 2006  M - Technology Denpasar - Indonesia"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Design By A.A.Ngr.Manik Artawan http://www.manikweb.net"
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   480
         Width           =   4455
      End
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      DataField       =   "Backup"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   1680
      Width           =   1910
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7320
      Top             =   1560
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000E&
      DataField       =   "name"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1200
      TabIndex        =   16
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000E&
      DataField       =   "Alamat"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1200
      TabIndex        =   15
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000E&
      DataField       =   "Telepon"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1200
      TabIndex        =   14
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H8000000E&
      DataField       =   "Fax"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   3240
      TabIndex        =   13
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H8000000E&
      DataField       =   "Email"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1200
      TabIndex        =   12
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H8000000E&
      DataField       =   "Logo"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1200
      TabIndex        =   11
      Top             =   2040
      Width           =   3255
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H8000000E&
      DataField       =   "Keterangan"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1200
      TabIndex        =   10
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton CmdCLose 
      Caption         =   "Close"
      Height          =   615
      Left            =   6000
      TabIndex        =   7
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton CMDCpuSys 
      Caption         =   "CPU Sys"
      Height          =   615
      Left            =   1200
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton CmdUpdate 
      Caption         =   "Update"
      Height          =   615
      Left            =   2400
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton CmdBackup 
      Caption         =   "Backup"
      Height          =   615
      Left            =   3600
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton CM02 
      Caption         =   "Load Logo"
      Height          =   615
      Left            =   4800
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      DataField       =   "Backup"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1320
      Width           =   1910
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      DataField       =   "Backup"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   2400
      Width           =   1910
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7320
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000009&
      Caption         =   "Design for :"
      Height          =   255
      Left            =   6960
      TabIndex        =   25
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Faximile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2520
      TabIndex        =   21
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Logo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LABEL MANAGER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1200
      TabIndex        =   18
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Notes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BACKUP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4800
      TabIndex        =   9
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "YOUR LOGO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   7080
      TabIndex        =   8
      Top             =   240
      Width           =   3015
   End
   Begin VB.Image Image7 
      Height          =   705
      Left            =   8760
      Picture         =   "SYSCONTROL.frx":0442
      Top             =   4080
      Width           =   1380
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   5160
      Picture         =   "SYSCONTROL.frx":0D8E
      Top             =   4440
      Width           =   2850
   End
   Begin VB.Image Image5 
      Height          =   660
      Left            =   4800
      Picture         =   "SYSCONTROL.frx":169A
      Top             =   600
      Width           =   1875
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Last backup created :"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4800
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Image Image4 
      Height          =   2175
      Left            =   7080
      Stretch         =   -1  'True
      Top             =   600
      Width           =   3015
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   -120
      Picture         =   "SYSCONTROL.frx":2A23
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   11160
   End
   Begin VB.Image Image1 
      Height          =   3945
      Left            =   -240
      Picture         =   "SYSCONTROL.frx":55E2
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   10785
   End
End
Attribute VB_Name = "SystemControl"
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
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Dim WindowsDir As String, NLoopsTimer As Byte, Interval As Date, IniTime As Date
Dim Source As String
Dim Destination As String

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Sub Form_Initialize()
   InitCommonControls
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
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
    Dim I As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For I = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, I, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function


'############################################3

Private Sub CM02_Click()
logoku
End Sub

Private Sub CmdBackup_Click()
If MsgBox("Are you sure to make backup file ?", vbYesNo + vbExclamation, "Confirmation") = vbYes Then
    Unload Register
    Source = "C:\VBDOCTOR\DataBase.mdb"
    Destination = "C:\VBDOCTOR\DataBackup.mdb"
    FileCopy Source, Destination
    MsgBox "File backup successfully"
    Register.Show
    Unload Me
    End If
End Sub

Private Sub CmdCLose_Click()
Unload Me
End Sub

Private Sub logoku()
Image4.Stretch = True
CommonDialog1.ShowOpen
Image4.Picture = LoadPicture(CommonDialog1.FileName)
Text6.Text = CommonDialog1.FileName
End Sub

Private Sub CMDCpuSys_Click()
Call StartSysInfo
End Sub

Private Sub CmdUpdate_Click()
Text9.Text = Text8.Text
With Adodc1.Recordset '
    !Name = Text1.Text
    !Alamat = Text2.Text
    !Telepon = Text3.Text
    !Email = Text5.Text
    !Keterangan = Text7.Text
    !Fax = Text4.Text
    !Logo = Text6.Text
    !Backup = Text9.Text
    .Update
    End With
    MsgBox "All system database already update, click ok to continue !"
End Sub

Private Sub Form_Load()
Image4.Stretch = True
Image4.Picture = LoadPicture(Text6.Text)
End Sub

Private Sub Timer1_Timer()
Text10.Text = Date & " " & Time
End Sub
