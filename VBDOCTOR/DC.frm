VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form DCPrint 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "DOCTOR CERTIFICAT"
   ClientHeight    =   8430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10395
   Icon            =   "DC.frx":0000
   LinkTopic       =   "Form17"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "General Setup"
      Height          =   4095
      Left            =   6000
      TabIndex        =   31
      Top             =   3240
      Width           =   3975
      Begin VB.CommandButton CMLanguage 
         Caption         =   "English"
         Height          =   375
         Left            =   120
         TabIndex        =   56
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton CM01 
         Caption         =   "Close"
         Height          =   375
         Left            =   3000
         TabIndex        =   55
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton CM03 
         Caption         =   "Print"
         Height          =   375
         Left            =   2040
         TabIndex        =   54
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton CM02 
         Caption         =   "Preview"
         Height          =   375
         Left            =   1080
         TabIndex        =   53
         Top             =   3480
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   1200
         TabIndex        =   51
         Top             =   2520
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   20643841
         CurrentDate     =   39020
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1200
         TabIndex        =   50
         Top             =   2160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   20643841
         CurrentDate     =   39020
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         DataSource      =   "Adodc2"
         Height          =   285
         Left            =   1200
         TabIndex        =   39
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         DataField       =   "Birth"
         DataSource      =   "Adodc2"
         Height          =   285
         Left            =   2280
         TabIndex        =   38
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         DataField       =   "Gender"
         DataSource      =   "Adodc2"
         Height          =   285
         Left            =   1200
         TabIndex        =   37
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         DataField       =   "MainAddress"
         DataSource      =   "Adodc2"
         Height          =   285
         Left            =   1200
         TabIndex        =   36
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   35
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox Lama 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   33
         Text            =   "Dr. Arya. S"
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   32
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000004&
         BorderWidth     =   2
         Height          =   3255
         Left            =   120
         Top             =   120
         Width           =   3735
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "For"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2640
         TabIndex        =   52
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "R.M"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   49
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "Days"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3360
         TabIndex        =   48
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   "Finish"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   47
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "Start"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   46
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   45
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   44
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl :"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1800
         TabIndex        =   43
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   42
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "Diagnose"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   41
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   40
         Top             =   1800
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   4560
      Top             =   7800
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\VBDOCTOR\DataBase.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\VBDOCTOR\DataBase.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from register"
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   6840
      Top             =   7800
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\VBDOCTOR\DataBase.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\VBDOCTOR\DataBase.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from medicalrecord"
      Caption         =   "Adodc2"
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
   Begin VB.TextBox Logo 
      DataField       =   "Logo"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2880
      TabIndex        =   30
      Top             =   7800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   7800
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      RecordSource    =   "select * from Label"
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
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "To whom it my concern"
      Height          =   255
      Left            =   3240
      TabIndex        =   59
      Top             =   2400
      Width           =   3975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Menerangkan dengan sebenarnya, pasien atas nama dibawah ini :"
      Height          =   255
      Left            =   1080
      TabIndex        =   58
      Top             =   2880
      Width           =   5775
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SURAT KETERANGAN SAKIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   57
      Top             =   2040
      Width           =   4095
   End
   Begin VB.Image Image2 
      DataField       =   "Logo"
      Height          =   1335
      Left            =   360
      Top             =   120
      Width           =   1815
   End
   Begin VB.Line Line4 
      BorderStyle     =   4  'Dash-Dot
      X1              =   240
      X2              =   10320
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line3 
      BorderStyle     =   4  'Dash-Dot
      DrawMode        =   7  'Invert
      X1              =   240
      X2              =   10560
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label33"
      DataField       =   "Keterangan"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   480
      TabIndex        =   29
      Top             =   1320
      Width           =   9735
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "Hari."
      Height          =   255
      Left            =   6960
      TabIndex        =   28
      Top             =   5160
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   7695
      Left            =   120
      Top             =   0
      Width           =   10095
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Label31"
      Height          =   255
      Left            =   2640
      TabIndex        =   27
      Top             =   4680
      Width           =   7575
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "Label30"
      Height          =   255
      Left            =   2640
      TabIndex        =   26
      Top             =   4320
      Width           =   7575
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "Label29"
      Height          =   255
      Left            =   2640
      TabIndex        =   25
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Label28"
      Height          =   255
      Left            =   2640
      TabIndex        =   24
      Top             =   3600
      Width           =   3735
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Label27"
      Height          =   255
      Left            =   2640
      TabIndex        =   23
      Top             =   3240
      Width           =   7575
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label22"
      Height          =   255
      Left            =   4080
      TabIndex        =   22
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      Height          =   255
      Left            =   960
      TabIndex        =   21
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label18"
      Height          =   255
      Left            =   6240
      TabIndex        =   20
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Label26"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   19
      Top             =   7320
      Width           =   8055
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Label25"
      Height          =   255
      Left            =   1920
      TabIndex        =   18
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Denpasar,"
      Height          =   255
      Left            =   960
      TabIndex        =   17
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Demikian Agar menjadikan maklum adanya."
      Height          =   255
      Left            =   960
      TabIndex        =   16
      Top             =   5760
      Width           =   6975
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Sampai dengan"
      Height          =   255
      Left            =   2640
      TabIndex        =   15
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Terhitung mulai tanggal"
      Height          =   255
      Left            =   7440
      TabIndex        =   14
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Memang benar dalam keadaan sakit dan membutuhkan istirahat selama"
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   5160
      Width           =   5175
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Diagnose       :"
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat           :"
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Kelamin         :"
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Lahir            :"
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama            :"
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   3240
      X2              =   7320
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   480
      X2              =   10080
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail :"
      Height          =   255
      Index           =   0
      Left            =   6480
      TabIndex        =   7
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Faximile :"
      Height          =   255
      Index           =   0
      Left            =   4440
      TabIndex        =   6
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Telepon :"
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   5
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      DataField       =   "Email"
      DataSource      =   "Adodc1"
      Height          =   255
      Index           =   0
      Left            =   7080
      TabIndex        =   4
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      DataField       =   "Fax"
      DataSource      =   "Adodc1"
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      DataField       =   "Telepon"
      DataSource      =   "Adodc1"
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      DataField       =   "Alamat"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   8895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      DataField       =   "name"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   8895
   End
End
Attribute VB_Name = "DCPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Sub Form_Initialize()
   InitCommonControls
End Sub

Private Sub CM01_Click()
Unload Me
End Sub

Private Sub CM02_Click()
Lama.Text = Val(DTPicker2.Value) - Val(DTPicker1.Value) + 1
Label27.Caption = Text1.Text
Label26.Caption = Text6.Text
Label20.Caption = DTPicker1.Value
Label22.Caption = DTPicker2.Value
Label18.Caption = Lama.Text
Label27.Caption = Text1.Text
Label28.Caption = Text2.Text
Label29.Caption = Text3.Text
Label30.Caption = Text4.Text
Label31.Caption = Text5.Text
Frame1.Visible = False
MsgBox "Temporary Preview", vbInformation
Frame1.Visible = True
End Sub

Private Sub CM03_Click()
If MsgBox("Yakin akan diprint ?", vbQuestion + vbYesNo + vbDefaultButton2, "Want to Exit ?") = vbYes Then
    Frame1.Visible = False
    DCPrint.PrintForm
    Frame1.Visible = True
    End If
End Sub



Private Sub CMLanguage_Click()
Label9.Caption = "I explain that the patient with the personal information data below :"
Label10.Caption = "DOCTOR CERTIFICATE"
Label12.Caption = "Name            :"
Label13.Caption = "Date of birth   :"
Label14.Caption = "Gender          :"
Label15.Caption = "Address         :"
Label17.Caption = "It is truly that, the patient name above need to take a nap at home for"
Label32.Caption = "Days"
Label19.Caption = "Start from date :"
Label21.Caption = "Until"
Label23.Caption = "Thank you for your attentions dan cooporation"
End Sub

Private Sub DTPicker1_Click()
Lama.Text = Val(DTPicker2.Value) - Val(DTPicker1.Value) + 1
End Sub

Private Sub DTPicker2_Click()
Lama.Text = Val(DTPicker2.Value) - Val(DTPicker1.Value) + 1
End Sub

Private Sub Form_Load()
Label25.Caption = Date
Image2.Stretch = True
Image2.Picture = LoadPicture(Logo.Text)
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Text7_Change()
With Adodc2
.RecordSource = "select * from medicalrecord where medicalrecord like '%" & _
Text7.Text & "%'"
.Refresh
End With
End Sub
