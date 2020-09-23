VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCard 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3240
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9810
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FRMCARD.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   9810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5040
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   1080
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7560
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      RecordSource    =   "Select * from Label"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   7560
      Top             =   1440
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      RecordSource    =   "select * from Medicalrecord"
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
   Begin VB.Shape Shape2 
      BorderStyle     =   3  'Dot
      Height          =   3135
      Left            =   4920
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   4815
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   3  'Dot
      Height          =   3135
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   4815
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label23"
      DataField       =   "Keterangan"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   5040
      TabIndex        =   18
      Top             =   720
      Width           =   4575
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label22"
      DataField       =   "Email"
      DataSource      =   "Adodc1"
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
      Height          =   255
      Left            =   5880
      TabIndex        =   17
      Top             =   1800
      Width           =   3735
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label19"
      DataField       =   "Fax"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   8400
      TabIndex        =   16
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label18"
      DataField       =   "Telepon"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   7200
      TabIndex        =   15
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label17"
      DataField       =   "Alamat"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   4920
      TabIndex        =   14
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This card is the clinic property. Its use is governed by the terms and conditions of clinic. If found, please return to clinic"
      Height          =   495
      Left            =   5040
      TabIndex        =   13
      Top             =   2520
      Width           =   4455
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1515
      Left            =   4920
      Picture         =   "FRMCARD.frx":000C
      Stretch         =   -1  'True
      Top             =   960
      Width           =   4800
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      DataField       =   "name"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   12
      Top             =   120
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label1b 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Pasien   :"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat Pasien  :"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      DataField       =   "Name"
      DataSource      =   "Adodc2"
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      DataField       =   "MainAddress"
      DataSource      =   "Adodc2"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      DataField       =   "Company"
      DataSource      =   "Adodc2"
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Roll            :"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      DataField       =   "Payroll"
      DataSource      =   "Adodc2"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Instansi            :"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      DataField       =   "Gender"
      DataSource      =   "Adodc2"
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Rekam Medik  :"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      DataField       =   "MedicalRecord"
      DataSource      =   "Adodc2"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "MEDICAL RECORD CARD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(KeyAscii As Integer)
Unload Me
End Sub

Private Sub Form_Load()
Image1.Picture = Register.Picture3.Image
End Sub

Private Sub Image2_Click()
Unload Me
End Sub

Private Sub Text1_Change()
With Adodc2
.RecordSource = "select * from medicalrecord where medicalrecord like '%" & _
Text1.Text & "%'"
.Refresh
End With
End Sub
