VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Register 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "VB DOCTOR 2007 - Copyright (C) Gung Manik - M-Technology"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11085
   Icon            =   "REGISTER.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   11085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4935
      Left            =   8880
      TabIndex        =   141
      Top             =   6480
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CommandButton BRCPrint 
         Caption         =   "Print"
         Height          =   375
         Left            =   240
         TabIndex        =   157
         Top             =   4320
         Width           =   855
      End
      Begin VB.CommandButton BRCImg 
         Caption         =   "Save Img"
         Height          =   375
         Left            =   1080
         TabIndex        =   156
         Top             =   4320
         Width           =   855
      End
      Begin VB.CommandButton BRCPView 
         Caption         =   "Preview"
         Height          =   375
         Left            =   1920
         TabIndex        =   155
         Top             =   4320
         Width           =   855
      End
      Begin VB.CommandButton BRCCLose 
         Caption         =   "Close"
         Height          =   375
         Left            =   2760
         TabIndex        =   154
         Top             =   4320
         Width           =   855
      End
      Begin VB.TextBox Text13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MedicalRecord"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   23
         TabIndex        =   143
         Text            =   "00.00.01"
         Top             =   240
         Width           =   1695
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   360
         ScaleHeight     =   1425
         ScaleWidth      =   3105
         TabIndex        =   142
         Top             =   720
         Width           =   3135
      End
      Begin MSComDlg.CommonDialog cd1 
         Left            =   3120
         Top             =   3720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         DataField       =   "Payroll"
         DataSource      =   "Adodc1"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1680
         TabIndex        =   153
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label54 
         BackStyle       =   0  'Transparent
         Caption         =   "Kelamin            :"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   152
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label53 
         BackStyle       =   0  'Transparent
         Caption         =   "ID CARD MAKER - EDITOR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   360
         TabIndex        =   151
         Top             =   2400
         Width           =   3255
      End
      Begin VB.Label Label52 
         BackStyle       =   0  'Transparent
         DataField       =   "MedicalRecord"
         DataSource      =   "Adodc1"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1680
         TabIndex        =   150
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label51 
         BackStyle       =   0  'Transparent
         Caption         =   "Rekam Medik  :"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   149
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         DataField       =   "Gender"
         DataSource      =   "Adodc1"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1680
         TabIndex        =   148
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Roll           :"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   147
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         DataField       =   "Name"
         DataSource      =   "Adodc1"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1680
         TabIndex        =   146
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label1b 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pasien   :"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   145
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         FillColor       =   &H8000000E&
         Height          =   4695
         Left            =   120
         Top             =   120
         Width           =   3615
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bar Code Text :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   360
         TabIndex        =   144
         Top             =   360
         Width           =   1380
      End
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   2880
      TabIndex        =   139
      Top             =   6720
      Width           =   975
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H8000000E&
      Caption         =   "Name Search"
      Height          =   255
      Left            =   2640
      TabIndex        =   138
      Top             =   6480
      Width           =   1335
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H8000000E&
      Caption         =   "Queue"
      Height          =   255
      Left            =   1560
      TabIndex        =   137
      Top             =   6720
      Width           =   855
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H8000000E&
      Caption         =   "Custom"
      Height          =   255
      Left            =   480
      TabIndex        =   136
      Top             =   6720
      Width           =   855
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H8000000E&
      Caption         =   "View All"
      Height          =   255
      Left            =   1560
      TabIndex        =   135
      Top             =   6480
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000009&
      Caption         =   "Today"
      Height          =   255
      Left            =   480
      TabIndex        =   134
      Top             =   6480
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   8400
      Top             =   7200
   End
   Begin VB.Frame FrameMR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   10440
      TabIndex        =   32
      Top             =   3960
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton Command3 
         Caption         =   "Close"
         Height          =   375
         Left            =   3240
         TabIndex        =   116
         Top             =   5040
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2280
         TabIndex        =   92
         Top             =   4680
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "New MR"
         Height          =   375
         Left            =   1200
         TabIndex        =   37
         Top             =   5040
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "< Reset"
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   5040
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   34
         Top             =   4680
         Width           =   1815
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "REGISTER.frx":0442
         Height          =   4095
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   7223
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "MEDICAL RECORDS"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "MedicalRecord"
            Caption         =   "M.Record"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Name"
            Caption         =   "Patient Name"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2355,024
            EndProperty
         EndProperty
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000004&
         BorderWidth     =   2
         Height          =   5415
         Left            =   120
         Top             =   120
         Width           =   4095
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Search M.R :"
         Height          =   255
         Left            =   2280
         TabIndex        =   93
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Search Name :"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   4440
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "MEDICAL RECORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   9360
      TabIndex        =   46
      Top             =   6000
      Visible         =   0   'False
      Width           =   6855
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   4800
         TabIndex        =   94
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox RMT02 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3720
         TabIndex        =   91
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton Frm3cls 
         Caption         =   "Close"
         Height          =   375
         Left            =   5520
         TabIndex        =   89
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton Frm3Update 
         Caption         =   "Update"
         Height          =   375
         Left            =   4560
         TabIndex        =   88
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton Frm3Save 
         Caption         =   "Save"
         Height          =   375
         Left            =   3600
         TabIndex        =   87
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox TXMR 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TxName 
         Height          =   285
         Left            =   1560
         TabIndex        =   65
         Top             =   720
         Width           =   4815
      End
      Begin VB.ComboBox TXGender 
         DataSource      =   "Adodc2"
         Height          =   315
         ItemData        =   "REGISTER.frx":0457
         Left            =   1560
         List            =   "REGISTER.frx":0461
         TabIndex        =   64
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox TXAge 
         DataSource      =   "Adodc2"
         Height          =   285
         Left            =   5760
         TabIndex        =   63
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox TxDate 
         Height          =   285
         Left            =   3600
         TabIndex        =   62
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox TxAddress 
         Height          =   285
         Left            =   1560
         TabIndex        =   61
         Top             =   2520
         Width           =   4815
      End
      Begin VB.TextBox TxAddress2 
         Height          =   285
         Left            =   1560
         TabIndex        =   60
         Top             =   2880
         Width           =   4815
      End
      Begin VB.TextBox TxPhone 
         Height          =   285
         Left            =   1560
         TabIndex        =   59
         Top             =   3240
         Width           =   1695
      End
      Begin VB.TextBox TXSelluler 
         Height          =   285
         Left            =   4320
         TabIndex        =   58
         Top             =   3240
         Width           =   2055
      End
      Begin VB.TextBox TxFax 
         Height          =   285
         Left            =   1560
         TabIndex        =   57
         Top             =   3600
         Width           =   1695
      End
      Begin VB.TextBox TXEmail 
         Height          =   285
         Left            =   4320
         TabIndex        =   56
         Top             =   3600
         Width           =   2055
      End
      Begin VB.TextBox TXMI 
         DataSource      =   "Adodc2"
         Height          =   285
         Left            =   1560
         TabIndex        =   55
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox TXNational 
         Height          =   285
         Left            =   3600
         TabIndex        =   54
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TxPayroll 
         Height          =   285
         Left            =   4320
         TabIndex        =   53
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox TxCorp 
         Height          =   285
         Left            =   1560
         TabIndex        =   52
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox TxJob 
         Height          =   285
         Left            =   4320
         TabIndex        =   51
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox TxEdu 
         Height          =   285
         Left            =   1560
         TabIndex        =   50
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox TxAlergy 
         Height          =   285
         Left            =   1560
         TabIndex        =   49
         Top             =   3960
         Width           =   4815
      End
      Begin VB.TextBox TXBlood 
         DataSource      =   "Adodc2"
         Height          =   285
         Left            =   5760
         TabIndex        =   48
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox TXNotes 
         Height          =   285
         Left            =   1560
         TabIndex        =   47
         Top             =   4320
         Width           =   4815
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000004&
         BorderWidth     =   2
         Height          =   5175
         Left            =   120
         Top             =   120
         Width           =   6615
      End
      Begin VB.Label LabelFrameMR 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "NONE"
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
         Left            =   1560
         TabIndex        =   90
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Medical Record"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   360
         TabIndex        =   86
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Name"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   360
         TabIndex        =   85
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   360
         TabIndex        =   84
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Ages"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   5160
         TabIndex        =   83
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   2640
         TabIndex        =   82
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Main Address"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   360
         TabIndex        =   81
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Temp  Address"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   360
         TabIndex        =   80
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   360
         TabIndex        =   79
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Seluler"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   3480
         TabIndex        =   78
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Faximile"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   360
         TabIndex        =   77
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   3480
         TabIndex        =   76
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Married Info"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   360
         TabIndex        =   75
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Nationality"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   2640
         TabIndex        =   74
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   3600
         TabIndex        =   73
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Coorporation"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   360
         TabIndex        =   72
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Job Info"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   3600
         TabIndex        =   71
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Educations "
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   360
         TabIndex        =   70
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Alergy "
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   360
         TabIndex        =   69
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Blood Type"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   4800
         TabIndex        =   68
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   360
         TabIndex        =   67
         Top             =   4320
         Width           =   1335
      End
   End
   Begin VB.Frame FrameRegister 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "REGISTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   10680
      TabIndex        =   1
      Top             =   2880
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton Command6 
         Caption         =   "New"
         Height          =   285
         Left            =   5040
         TabIndex        =   158
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton FrmSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   2760
         TabIndex        =   30
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton FrmUpdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   3720
         TabIndex        =   27
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton FrmClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   4680
         TabIndex        =   26
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox Txtreg 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TxtTime 
         Height          =   285
         Left            =   3960
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox TxtMR 
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TxtName 
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Top             =   1320
         Width           =   3975
      End
      Begin VB.TextBox TxtAge 
         Height          =   285
         Left            =   3120
         TabIndex        =   9
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox TxtDiagnose 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   2040
         Width           =   3975
      End
      Begin VB.TextBox TxtPayroll 
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox TxtDoctor 
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   2400
         Width           =   3975
      End
      Begin VB.TextBox TxtCase 
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   2760
         Width           =   3975
      End
      Begin VB.TextBox TxtTheraphy 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   3120
         Width           =   3975
      End
      Begin VB.TextBox TxtDate 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox TxtGender 
         Height          =   315
         ItemData        =   "REGISTER.frx":0473
         Left            =   4560
         List            =   "REGISTER.frx":047D
         TabIndex        =   2
         Top             =   1680
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000004&
         BorderWidth     =   2
         Height          =   4215
         Left            =   120
         Top             =   120
         Width           =   5775
      End
      Begin VB.Label LabelFrameRegister 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "NONE"
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
         Left            =   1560
         TabIndex        =   29
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Register"
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   480
         TabIndex        =   25
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Date In"
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   480
         TabIndex        =   24
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Time In"
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   2640
         TabIndex        =   23
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Medical Record"
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   2640
         TabIndex        =   22
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Name"
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   480
         TabIndex        =   21
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   3840
         TabIndex        =   20
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Ages"
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   2640
         TabIndex        =   19
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Diagnose"
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll"
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor Name"
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Medical Case"
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Theraphy"
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   3120
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4815
      Left            =   9720
      TabIndex        =   117
      Top             =   5400
      Visible         =   0   'False
      Width           =   5655
      Begin VB.TextBox Text11 
         DataField       =   "Expense"
         DataSource      =   "Adodc4"
         Height          =   285
         Left            =   2640
         TabIndex        =   131
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text10 
         DataField       =   "ItemsName"
         DataSource      =   "Adodc4"
         Height          =   285
         Left            =   720
         TabIndex        =   130
         Top             =   1200
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton CMDFR4refresh 
         Caption         =   "View All"
         Height          =   375
         Left            =   1920
         TabIndex        =   129
         Top             =   3960
         Width           =   855
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4560
         TabIndex        =   128
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "REGISTER.frx":048F
         Left            =   3480
         List            =   "REGISTER.frx":049C
         TabIndex        =   127
         Text            =   "Medicine"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2640
         TabIndex        =   126
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text7 
         Height          =   315
         Left            =   720
         TabIndex        =   125
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton CMDFR4Add 
         Caption         =   "Add"
         Height          =   375
         Left            =   2760
         TabIndex        =   124
         Top             =   3960
         Width           =   855
      End
      Begin VB.CommandButton CMDFR4Cls 
         Caption         =   "Close"
         Height          =   375
         Left            =   4440
         TabIndex        =   123
         Top             =   3960
         Width           =   855
      End
      Begin VB.CommandButton CMDFR4Snd 
         Caption         =   "Send"
         Height          =   375
         Left            =   3600
         TabIndex        =   122
         Top             =   3960
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   3600
         TabIndex        =   120
         Top             =   3480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "REGISTER.frx":04C4
         Left            =   360
         List            =   "REGISTER.frx":04D1
         TabIndex        =   119
         Text            =   "Medicine"
         Top             =   4200
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "REGISTER.frx":04F9
         Height          =   3255
         Left            =   360
         TabIndex        =   118
         Top             =   600
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   5741
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "ItemsName"
            Caption         =   "Items Name"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Expense"
            Caption         =   "Expense"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Catagory"
            Caption         =   "Catagory"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Stock"
            Caption         =   "Stock"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1950,236
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   824,882
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   585,071
            EndProperty
         EndProperty
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "Index Catagory :"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   121
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H80000004&
         BorderWidth     =   2
         Height          =   4575
         Left            =   120
         Top             =   120
         Width           =   5415
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   5415
      Left            =   10200
      TabIndex        =   97
      Top             =   4800
      Visible         =   0   'False
      Width           =   7215
      Begin VB.CommandButton Command4 
         Caption         =   "V"
         Height          =   285
         Left            =   480
         TabIndex        =   132
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox TxS07 
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   110
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox TxS06 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   109
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Frm2Calc 
         Caption         =   "Recalc"
         Height          =   375
         Left            =   4920
         TabIndex        =   108
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox TxS05 
         Height          =   285
         Left            =   5520
         TabIndex        =   107
         Text            =   "0"
         Top             =   840
         Width           =   1000
      End
      Begin VB.TextBox TxS04 
         Height          =   285
         Left            =   4680
         TabIndex        =   106
         Text            =   "0"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox TxS03 
         Height          =   285
         Left            =   3600
         TabIndex        =   105
         Text            =   "0"
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Frm2Update 
         Caption         =   "Save"
         Height          =   375
         Left            =   3960
         TabIndex        =   104
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton Frm2Close 
         Caption         =   "Close"
         Height          =   375
         Left            =   5880
         TabIndex        =   103
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox TxS02 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3120
         TabIndex        =   101
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox TxS01 
         Height          =   285
         Left            =   840
         TabIndex        =   100
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox TxS08 
         Height          =   285
         Left            =   1200
         TabIndex        =   99
         Text            =   "0"
         Top             =   4680
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "REGISTER.frx":050E
         Height          =   3375
         Left            =   480
         TabIndex        =   98
         Top             =   1200
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "Service"
            Caption         =   "Service"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Sum"
            Caption         =   "Sum"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Price"
            Caption         =   "Price"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Discont"
            Caption         =   "Discont"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Total"
            Caption         =   "Total"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   2355,024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   480,189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   810,142
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1035,213
            EndProperty
         EndProperty
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Patinet Name :"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   840
         TabIndex        =   111
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   480
         TabIndex        =   102
         Top             =   4680
         Width           =   615
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H80000004&
         BorderWidth     =   2
         Height          =   5175
         Left            =   120
         Top             =   120
         Width           =   6975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000010&
      Height          =   3015
      Left            =   480
      TabIndex        =   38
      Top             =   2760
      Visible         =   0   'False
      Width           =   4150
      Begin VB.CommandButton CmdFrame1cls 
         Caption         =   "Close"
         Height          =   375
         Left            =   3120
         TabIndex        =   40
         Top             =   2520
         Width           =   855
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2310
         Left            =   120
         TabIndex        =   39
         Top             =   120
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   4075
         _Version        =   393216
         ForeColor       =   -2147483639
         BackColor       =   -2147483634
         Appearance      =   0
         MonthBackColor  =   -2147483625
         StartOfWeek     =   58327042
         CurrentDate     =   39047
      End
   End
   Begin VB.CommandButton CMDMisc 
      Caption         =   "Add Items"
      Height          =   615
      Left            =   7440
      TabIndex        =   115
      ToolTipText     =   "Add new items"
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox TEXTMR 
      Alignment       =   1  'Right Justify
      DataField       =   "MedicalRecord"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   480
      TabIndex        =   114
      Text            =   "MR"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CMDMenu 
      Caption         =   "My Menu"
      Height          =   615
      Left            =   6240
      TabIndex        =   113
      ToolTipText     =   "Data Editor"
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox XPrint 
      DataField       =   "Register"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   480
      TabIndex        =   112
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Ani 
      Height          =   285
      Left            =   10560
      TabIndex        =   96
      Text            =   "0"
      Top             =   6240
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7920
      Top             =   7200
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1680
      TabIndex        =   95
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   240
      ScaleHeight     =   585
      ScaleWidth      =   10545
      TabIndex        =   41
      Top             =   120
      Width           =   10575
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   4560
         ScaleHeight     =   390
         ScaleWidth      =   1665
         TabIndex        =   43
         Top             =   70
         Width           =   1695
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Record (s) :"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   60
            Width           =   855
         End
         Begin VB.Label RecordLabel 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   255
            Left            =   960
            TabIndex        =   44
            Top             =   60
            Width           =   615
         End
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         DataField       =   "Name"
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
         Height          =   420
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   70
         Width           =   3615
      End
      Begin VB.Image Image10 
         Height          =   330
         Left            =   2640
         Picture         =   "REGISTER.frx":0523
         ToolTipText     =   "Quit"
         Top             =   120
         Width           =   360
      End
      Begin VB.Image Image9 
         Height          =   330
         Left            =   3120
         Picture         =   "REGISTER.frx":06AD
         ToolTipText     =   "Control Panels"
         Top             =   120
         Width           =   360
      End
      Begin VB.Image Image8 
         Height          =   330
         Left            =   3600
         Picture         =   "REGISTER.frx":0837
         ToolTipText     =   "Show Clock"
         Top             =   120
         Width           =   360
      End
      Begin VB.Image Image7 
         Height          =   360
         Left            =   4080
         Picture         =   "REGISTER.frx":09C1
         ToolTipText     =   "Hide / Show Data Grid"
         Top             =   120
         Width           =   360
      End
      Begin VB.Image Image6 
         Height          =   345
         Left            =   120
         Picture         =   "REGISTER.frx":0ED2
         Top             =   120
         Width           =   2250
      End
      Begin VB.Image Image2 
         Height          =   360
         Left            =   6360
         Picture         =   "REGISTER.frx":4B62
         Top             =   105
         Width           =   360
      End
   End
   Begin VB.TextBox FilterTanggal 
      Height          =   285
      Left            =   480
      TabIndex        =   31
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CMDAdd 
      Caption         =   "Register"
      Height          =   615
      Left            =   5040
      TabIndex        =   28
      ToolTipText     =   "Add New Register"
      Top             =   6360
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "REGISTER.frx":501A
      Height          =   5415
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   9551
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "PATIENTS LIST"
      ColumnCount     =   14
      BeginProperty Column00 
         DataField       =   "Register"
         Caption         =   "REG"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "DateIn"
         Caption         =   "DATE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "TimeIn"
         Caption         =   "TIME"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "MedicalRecord"
         Caption         =   "M.RCD"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Name"
         Caption         =   "NAME"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Gender"
         Caption         =   "GENDER"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Age"
         Caption         =   "AGES"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Diagnose"
         Caption         =   "DIAGNOSE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Payroll"
         Caption         =   "PAYROLL"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Doctor"
         Caption         =   "DOCTOR"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "Case"
         Caption         =   "CASE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "Theraphy"
         Caption         =   "THERAPHY"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "Queue"
         Caption         =   "QUEUE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "Expense"
         Caption         =   "EXPENSE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   540,284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   975,118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   659,906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2684,977
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   824,882
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   540,284
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1830,047
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   854,929
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1454,74
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1260,284
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   705,26
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1365,165
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   6000
      Top             =   7320
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "select * from items"
      Caption         =   "Adodc4"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   4080
      Top             =   7320
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "select * from Service"
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
      Left            =   240
      Top             =   7320
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "select * from MedicalRecord"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2160
      Top             =   7320
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
   Begin VB.Frame Frame5 
      BackColor       =   &H80000009&
      Caption         =   "Filters"
      Height          =   855
      Left            =   240
      TabIndex        =   140
      Top             =   6240
      Width           =   3855
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000004&
      Height          =   255
      Left            =   9120
      TabIndex        =   133
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   -120
      Picture         =   "REGISTER.frx":502F
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   11520
   End
   Begin VB.Image Image1 
      Height          =   6225
      Left            =   0
      Picture         =   "REGISTER.frx":7BEE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11160
   End
   Begin VB.Image Image3 
      Height          =   705
      Left            =   11640
      Picture         =   "REGISTER.frx":8CEF
      Top             =   6240
      Width           =   1380
   End
   Begin VB.Image Image4 
      Height          =   705
      Left            =   9480
      Picture         =   "REGISTER.frx":C602
      ToolTipText     =   "Click it for more information and support"
      Top             =   6240
      Width           =   1380
   End
   Begin VB.Menu MENU 
      Caption         =   "MENU"
      Visible         =   0   'False
      Begin VB.Menu RegDetails 
         Caption         =   "Register Detail"
      End
      Begin VB.Menu MedicalRecord 
         Caption         =   "Medical Record"
      End
      Begin VB.Menu Service 
         Caption         =   "Add Service"
      End
      Begin VB.Menu Checkout 
         Caption         =   "Checkout"
      End
      Begin VB.Menu PrintReceipt 
         Caption         =   "Print Receipt"
      End
      Begin VB.Menu DrCertificate 
         Caption         =   "Dr Certificate"
      End
      Begin VB.Menu IDCard 
         Caption         =   "ID Card Maker"
      End
      Begin VB.Menu Blank 
         Caption         =   "-----------------"
      End
      Begin VB.Menu Delete 
         Caption         =   "Delete It"
      End
   End
End
Attribute VB_Name = "Register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Sub Form_Initialize()
   InitCommonControls
End Sub

Private Sub BRCCLose_Click()
Frame6.Visible = False
frmCard.Hide
End Sub

Private Sub BRCImg_Click()
cd1.FilterIndex = 1
    On Error GoTo ErrHandler
    cd1.FileName = Picture3.Name
    cd1.CancelError = True
    cd1.Flags = cdlOFNOverwritePrompt + cdlOFNNoReadOnlyReturn
    cd1.Filter = "Bitmaps (*.bmp)|*.bmp"
    cd1.ShowSave
    DoEvents
    Picture3.Picture = Picture3.Image
    Select Case cd1.FilterIndex
        Case 1:
        SavePicture Picture3.Picture, cd1.FileName
    End Select
    Exit Sub
ErrHandler:
        Exit Sub
End Sub

Private Sub BRCPrint_Click()
SetGambar
frmCard.Text1.Text = Text13.Text
frmCard.PrintForm
End Sub

Private Sub SetGambar()
cd1.FilterIndex = 1
cd1.FileName = "C:\VBDOCTOR\gambar.bmp"
DoEvents
Picture3.Picture = Picture3.Image
SavePicture Picture3.Picture, cd1.FileName
End Sub

Private Sub BRCPView_Click()
frmCard.Show
frmCard.Text1.Text = Text13.Text
End Sub

Private Sub Checkout_Click()
If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
With Adodc1.Recordset
!Queue = "Out"
.Update
End With
Text4.Text = "": Text5.Text = "": Text5.Text = "IN"
Option4.Value = True
End Sub

Private Sub CMDAdd_Click()
FrameRegister.Visible = True
frameregistercls
REGCOUNTERLOAD
LabelFrameRegister.Caption = "ADD MODE"
FrmSave.Enabled = True
FrmUpdate.Enabled = False
Text1.Text = "": Text3.Text = "": Text4.Text = ""
TxtMR.SetFocus
End Sub

Private Sub CMDDelete_Click()
With Adodc1.Recordset
If .RecordCount = 0 Then
    Else
    .Delete
    RecordLabel.Caption = .RecordCount
    End If
End With
End Sub


Private Sub CMDFR4Add_Click()
If Text7.Text = "" Then Exit Sub
With Adodc4.Recordset
    .AddNew
    !ItemsName = Text7.Text
    !Expense = Text8.Text
    !Catagory = Combo2.Text
    !Stock = Text9.Text
    .Update
    End With
End Sub

Private Sub CMDFR4Cls_Click()
Frame4.Visible = False
End Sub

Private Sub CMDFR4refresh_Click()
Text6.Text = ""
End Sub

Private Sub CMDFR4Snd_Click()
If Frame2.Visible = False Then Exit Sub
TxS01.Text = Text10.Text
TxS03.Text = Text11.Text
End Sub

Private Sub CmdFrame1cls_Click()
Frame1.Visible = False
End Sub

Private Sub CMDMenu_Click()
PopupMenu MENU
End Sub

Private Sub CMDMisc_Click()
Frame4.Visible = True
End Sub


Private Sub Combo1_Click()
Text6.Text = Combo1.Text
End Sub

Private Sub Command1_Click()
RMT02.Text = ""
End Sub

Private Sub Command2_Click()
FrameMR.Visible = False
Frame3.Visible = True
RMCLear
LabelFrameMR.Caption = "ADD MODE"
Frm3Update.Enabled = False
MRCOUNTERLOAD
RMTSet
TXMR.Text = RMT02.Text
End Sub

Private Sub RMCLear()
    Text3.Text = ""
    TXMR.Text = ""
    TxName.Text = ""
    TXMI.Text = ""
    TXNational.Text = ""
    TXBlood.Text = ""
    TXGender.Text = ""
    TxDate.Text = ""
    TXAge.Text = ""
    TxCorp.Text = ""
    TxPayroll.Text = ""
    TxEdu.Text = ""
    TxJob.Text = ""
    TxAddress.Text = ""
    TxAddress2.Text = ""
    TxPhone.Text = ""
    TXSelluler.Text = ""
    TxFax.Text = ""
    TXEmail.Text = ""
    TxAlergy.Text = ""
    TXNotes.Text = ""
End Sub


Private Sub Command3_Click()
FrameMR.Visible = False
End Sub



Private Sub Command4_Click()
Frame4.Visible = True
End Sub


Private Sub Command6_Click()
Command2_Click
End Sub

Private Sub DataGrid1_mousedown(button As Integer, sift As Integer, X As Single, Y As Single)
If button = 2 Then
PopupMenu MENU
End If
End Sub

Private Sub DataGrid2_Click()
TxtMR.Text = Adodc2.Recordset!MedicalRecord
TxtName.Text = Adodc2.Recordset!Name
TxtAge.Text = Adodc2.Recordset!Age
TxtGender.Text = Adodc2.Recordset!Gender
TxtPayroll.Text = Adodc2.Recordset!payroll
FrameMR.Visible = False
End Sub

Private Sub Delete_Click()
If MsgBox("Are you sure to delete ?", vbYesNo + vbExclamation, "Confirmation") = vbYes Then
With Adodc1.Recordset
    If .RecordCount = 0 Then
        Else
        .Delete
        RecordLabel.Caption = .RecordCount
        End If
    End With
End If
End Sub

Private Sub DrCertificate_Click()
DCPrint.Show
DCPrint.Text7.Text = TEXTMR.Text
End Sub

Private Sub FilterTanggal_Change()
With Adodc1
.RecordSource = "select * from register where DateIn like '%" & _
FilterTanggal.Text & "%'"
.Refresh
End With
RecordLabel.Caption = Adodc1.Recordset.RecordCount
End Sub

Private Sub Form_Load()
FilterTanggal.Text = Date
Text5.Text = ""
RecordLabel.Caption = Adodc1.Recordset.RecordCount
Text6.Text = Combo1.Text
SettingFrame
Call DrawBarcode(Text13, Picture3)
End Sub

Private Sub SettingFrame()
Frame1.Top = 2760
Frame1.Left = 480
Frame2.Top = 720
Frame2.Left = 2040
Frame3.Top = 840
Frame3.Left = 3000
Frame4.Top = 840
Frame4.Left = 3000
FrameRegister.Top = 1320
FrameRegister.Left = 3000
FrameMR.Top = 840
FrameMR.Left = 6240
End Sub

Private Sub Frm2Calc_Click()
hitung03
End Sub

Private Sub hitung03()
TxS08.Text = 0
Adodc3.Refresh
With Adodc3.Recordset
    If .RecordCount = 0 Then
        Exit Sub
        End If
    If .RecordCount > 0 Then
        .MoveFirst
        End If
    Do While Not .EOF
        If TxS06.Text = !Register Then
        a = !Total
        TxS08.Text = Val(TxS08.Text) + a
        End If
        .MoveNext
    Loop
    End With
    
With Adodc1.Recordset
    !Expense = TxS08.Text
    .Update
    End With
End Sub

Private Sub Frm2Close_Click()
Frame2.Visible = False
End Sub

Private Sub Frm2Update_Click()
If TxS01.Text = "" Then Exit Sub
hitungx01
With Adodc3.Recordset
    .AddNew
    !Register = TxS06.Text
    !Service = TxS01.Text
    !Sum = TxS02.Text
    !Price = TxS03.Text
    !Discont = TxS04.Text
    !Total = TxS05.Text
    .Update
    End With
hitung03
End Sub

Private Sub Frm3cls_Click()
Frame3.Visible = False
End Sub

Private Sub Frm3Save_Click()
Frame3.Visible = False
    TxtName.Text = TxName.Text
    TxtMR.Text = TXMR.Text
    TxtGender.Text = TXGender.Text
    TxtPayroll.Text = TxPayroll.Text
    TxtAge.Text = TXAge.Text
With Adodc2.Recordset
    .AddNew
    !MedicalRecord = TXMR.Text
    !Name = TxName.Text
    !MariedStatue = TXMI.Text
    !Country = TXMI.Text
    !BllodType = TXBlood.Text
    !Gender = TXGender.Text
    !Birth = TxDate.Text
    !Age = TXAge.Text
    !Company = TxCorp.Text
    !payroll = TxPayroll.Text
    !Education = TxEdu.Text
    !Job = TxJob.Text
    !MainAddress = TxAddress.Text
    !TempAddress = TxAddress2.Text
    !Phone = TxPhone.Text
    !Selluler = TXSelluler.Text
    !Faximile = TxFax.Text
    !Email = TXEmail.Text
    !Alergic = TxAlergy.Text
    !Notes = TXNotes.Text
    .Update
    End With
    MRCOUNTERSAVE
End Sub

Private Sub Frm3Update_Click()
Frame3.Visible = False
With Adodc2.Recordset
    !MedicalRecord = TXMR.Text
    !Name = TxName.Text
    !MariedStatue = TXMI.Text
    !Country = TXMI.Text
    !BllodType = TXBlood.Text
    !Gender = TXGender.Text
    !Birth = TxDate.Text
    !Age = TXAge.Text
    !Company = TxCorp.Text
    !payroll = TxPayroll.Text
    !Education = TxEdu.Text
    !Job = TxJob.Text
    !MainAddress = TxAddress.Text
    !TempAddress = TxAddress2.Text
    !Phone = TxPhone.Text
    !Selluler = TXSelluler.Text
    !Faximile = TxFax.Text
    !Email = TXEmail.Text
    !Alergic = TxAlergy.Text
    !Notes = TXNotes.Text
    .Update
    End With
End Sub

Private Sub FrmClose_Click()
FrameRegister.Visible = False
FrameMR.Visible = False
End Sub

Private Sub TodayData()
Option1.Value = True
End Sub

Private Sub FrmSave_Click()
TodayData
If TxtMR.Text = "" Then
    FrameRegister.Visible = False
    Exit Sub
    End If
With Adodc1.Recordset
.AddNew
!Register = Txtreg.Text
!MedicalRecord = TxtMR.Text
!DateIn = TxtDate.Text
!TimeIn = TxtTime.Text
!Gender = TxtGender.Text
!Name = TxtName.Text
!payroll = TxtPayroll.Text
!Age = TxtAge.Text
!Diagnose = TxtDiagnose.Text
!Doctor = TxtDoctor.Text
!Case = TxtCase.Text
!Queue = "IN"
!Theraphy = TxtTheraphy.Text
.Update
RecordLabel.Caption = .RecordCount
End With
REGCOUNTERSAVE
REGCOUNTERLOAD
MsgBox "Saving successful...", vbInformation
End Sub

Private Sub FrmUpdate_Click()
FrameRegister.Visible = False
If LabelFrameRegister = "Add Mode" Then Exit Sub
With Adodc1.Recordset
If .RecordCount = 0 Then
    Else
    With Adodc1.Recordset
    If .RecordCount = 0 Then Exit Sub
     !Register = Txtreg.Text
     !MedicalRecord = TxtMR.Text
     !DateIn = TxtDate.Text
     !TimeIn = TxtTime.Text
     !Gender = TxtGender.Text
     !Name = TxtName.Text
     !payroll = TxtPayroll.Text
     !Age = TxtAge.Text
     !Diagnose = TxtDiagnose.Text
     !Doctor = TxtDoctor.Text
     !Case = TxtCase.Text
     !Theraphy = TxtTheraphy.Text
     .Update
    End With
    End If
End With
End Sub



Private Sub IDCard_Click()
Frame6.Visible = True
End Sub

Private Sub Image10_Click()
If MsgBox("Are you sure to exit ?", vbYesNo + vbExclamation, "Confirmation") = vbYes Then
End
End If
End Sub

Private Sub Image7_Click()
If DataGrid1.Visible = False Then DataGrid1.Visible = True: Exit Sub
If DataGrid1.Visible = True Then DataGrid1.Visible = False: Exit Sub
End Sub

Private Sub Image8_Click()
AnalocClock.Show
End Sub

Private Sub Image9_Click()
SystemControl.Show
End Sub

Private Sub MedicalRecord_Click()
Frame3.Visible = True
LabelFrameMR.Caption = "EDIT MODE"
Frm3Save.Enabled = False
Frm3Update.Enabled = True
If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
Text4.Text = Adodc1.Recordset!MedicalRecord
With Adodc2.Recordset
    If .RecordCount = 0 Then Exit Sub
    Text3.Text = Text4.Text
    TXMR.Text = !MedicalRecord
    TxName.Text = !Name
    TXMI.Text = !MariedStatue
    TXNational.Text = !Country
    TXBlood.Text = !BllodType
    TXGender.Text = !Gender
    TxDate.Text = !Birth
    TXAge.Text = !Age
    TxCorp.Text = !Company
    TxPayroll.Text = !payroll
    TxEdu.Text = !Education
    TxJob.Text = !Job
    TxAddress.Text = !MainAddress
    TxAddress2.Text = !TempAddress
    TxPhone.Text = !Phone
    TXSelluler.Text = !Selluler
    TxFax.Text = !Faximile
    TXEmail.Text = !Email
    TxAlergy.Text = !Alergic
    TXNotes.Text = !Notes
End With
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
FilterTanggal.Text = MonthView1.Value
End Sub

Private Sub Option1_Click()
Text5.Text = ""
FilterTanggal.Text = Date
End Sub

Private Sub Option2_Click()
FilterTanggal.Text = ""
Text5.Text = ""
End Sub

Private Sub Option3_Click()
Frame1.Visible = True
Text5.Text = ""
End Sub

Private Sub Option4_Click()
FilterTanggal.Text = ""
Text5.Text = "IN"
End Sub

Private Sub Option5_Click()
Text12.SetFocus
End Sub

Private Sub PrintReceipt_Click()
With DE.rsRegister
    .Filter = "Register='" & XPrint.Text & "'"
End With
DE.rsRegister.Open
DataReport1.Refresh
DataReport1.LeftMargin = 1
DataReport1.TopMargin = 1
DataReport1.Show
DE.rsRegister.Close
End Sub

Private Sub RegDetails_Click()
LabelFrameRegister = "EDIT MODE"
FrmSave.Enabled = False
FrmUpdate.Enabled = True
FrameRegister.Visible = True
With Adodc1.Recordset
    If .RecordCount = 0 Then Exit Sub
    Txtreg.Text = !Register
    TxtMR.Text = !MedicalRecord
    TxtDate.Text = !DateIn
    TxtTime.Text = !TimeIn
    TxtGender.Text = !Gender
    TxtName.Text = !Name
    TxtPayroll.Text = !payroll
    TxtAge.Text = !Age
    TxtDiagnose.Text = !Diagnose
    TxtDoctor.Text = !Doctor
    TxtCase.Text = !Case
    TxtTheraphy.Text = !Theraphy
End With
End Sub

Private Sub frameregistercls()
    Txtreg.Text = "0"
    TxtMR.Text = ""
    TxtDate.Text = Date
    TxtTime.Text = Time
    TxtGender.Text = ""
    TxtName.Text = ""
    TxtPayroll.Text = ""
    TxtAge.Text = ""
    TxtDiagnose.Text = ""
    TxtDoctor.Text = "Dr.Arya S"
    TxtCase.Text = ""
    TxtTheraphy.Text = ""
End Sub

Private Sub RMTSet()
If Len(Trim(RMT02.Text)) > 5 Then RMT02.Text = RMT02.Text: Exit Sub
If Len(Trim(RMT02.Text)) > 4 Then RMT02.Text = "0" & RMT02.Text: Exit Sub
If Len(Trim(RMT02.Text)) > 3 Then RMT02.Text = "00" & RMT02.Text: Exit Sub
If Len(Trim(RMT02.Text)) > 2 Then RMT02.Text = "000" & RMT02.Text: Exit Sub
If Len(Trim(RMT02.Text)) > 1 Then RMT02.Text = "0000" & RMT02.Text: Exit Sub
If Len(Trim(RMT02.Text)) > 0 Then RMT02.Text = "00000" & RMT02.Text: Exit Sub
End Sub

Private Sub Service_Click()
If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
Frame2.Visible = True
TxS01.Text = "": TxS02.Text = "1": TxS03.Text = "0"
TxS06.Text = Adodc1.Recordset!Register
TxS07.Text = Adodc1.Recordset!Name
TxS02.Text = "1"
Frm2Calc_Click
TxS01.SetFocus
End Sub

Private Sub Text1_Change()
With Adodc2
.RecordSource = "select * from medicalrecord where name like '%" & _
Text1.Text & "%'"
.Refresh
End With
End Sub

Private Sub Text12_Change()
Option5.Value = True
With Adodc1
.RecordSource = "select * from Register where name like '%" & _
Text12.Text & "%'"
.Refresh
End With
End Sub

Private Sub Text13_Change()
Call DrawBarcode(Text13, Picture3)
End Sub

Private Sub Text3_Change()
With Adodc2
.RecordSource = "select * from medicalrecord where MedicalREcord like '%" & _
Text3.Text & "%'"
.Refresh
End With
End Sub

Private Sub Text4_Change()
With Adodc2
.RecordSource = "select * from medicalrecord where MedicalRecord like '%" & _
Text4.Text & "%'"
.Refresh
End With
End Sub

Private Sub Text5_Change()
With Adodc1
.RecordSource = "select * from register where Queue like '%" & _
Text5.Text & "%'"
.Refresh
End With
RecordLabel.Caption = Adodc1.Recordset.RecordCount
End Sub


Private Sub Text6_Change()
With Adodc4
.RecordSource = "select * from Items where Catagory like '%" & _
Text6.Text & "%'"
.Refresh
End With
End Sub

Private Sub Timer1_Timer()
If Ani.Text = "0" Then Image3.Left = 9480: Ani.Text = "1": Exit Sub
If Ani.Text = "1" Then Image3.Left = 11460: Ani.Text = "0"
End Sub

Private Sub Timer2_Timer()
Label39.Caption = Date & " , " & Time
End Sub


Private Sub TxS02_Validate(Cancel As Boolean)
If Not IsNumeric(TxS02.Text) Then
    MsgBox "Please enter numbers only.", vbInformation
    TxS02.Text = ""
    Cancel = True
    Else
    hitungx01
End If
End Sub

Private Sub hitungx01()
TxS05.Text = (TxS02.Text * TxS03) - ((TxS02.Text * TxS03) * (TxS04.Text / 100))
End Sub

Private Sub TxS03_Validate(Cancel As Boolean)
If Not IsNumeric(TxS03.Text) Then
    MsgBox "Please enter numbers only.", vbInformation
    TxS03.Text = ""
    Cancel = True
    Else
    hitungx01
End If
End Sub

Private Sub TxS04_Validate(Cancel As Boolean)
If Not IsNumeric(TxS04.Text) Then
    MsgBox "Please enter numbers only.", vbInformation
    TxS04.Text = "0"
    Cancel = True
    Else
    hitungx01
End If
End Sub

Private Sub TxS08_Validate(Cancel As Boolean)
If Not IsNumeric(TxS08.Text) Then
    MsgBox "Please enter numbers only.", vbInformation
    TxS08.Text = "0"
    Cancel = True
    Else
End If
End Sub

Private Sub TxS06_Change()
With Adodc3
.RecordSource = "select * from Service where register like '%" & _
TxS06.Text & "%'"
.Refresh
End With
End Sub

Private Sub TxtMR_Click()
If LabelFrameRegister = "EDIT MODE" Then Exit Sub
FrameMR.Visible = True
End Sub

Private Sub REGCOUNTERSAVE()
    Dim intFileHandle As Integer
    Dim strRETP As String
    strRETP = "Hi There"
    intFileHandle = FreeFile
    Open "COUNTER.txt" For Output As #intFileHandle
    Print #intFileHandle, Txtreg.Text
    Close #intFileHandle
End Sub

Private Sub REGCOUNTERLOAD()
    Dim intFileHandle As Integer
    Dim strRETP As String
    intFileHandle = FreeFile
    Open "COUNTER.txt" For Input As #intFileHandle
    Line Input #intFileHandle, strRETP
    Close #intFileHandle
    Txtreg.Text = strRETP + 1
End Sub

Private Sub MRCOUNTERLOAD()
    Dim intFileHandle2 As Integer
    Dim strRETP2 As String
    intFileHandle2 = FreeFile
    Open "COUNTER2.txt" For Input As #intFileHandle2
    Line Input #intFileHandle2, strRETP2
    Close #intFileHandle2
    RMT02.Text = strRETP2 + 1
End Sub

Private Sub MRCOUNTERSAVE()
    Dim intFileHandle2 As Integer
    Dim strRETP2 As String
    strRETP2 = "Hi There"
    intFileHandle2 = FreeFile
    Open "COUNTER2.txt" For Output As #intFileHandle2
    Print #intFileHandle2, TXMR.Text
    Close #intFileHandle2
End Sub


