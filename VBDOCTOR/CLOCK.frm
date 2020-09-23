VERSION 5.00
Begin VB.Form AnalocClock 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Analog Clock"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   2460
   Icon            =   "CLOCK.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "CLOCK.frx":0442
   ScaleHeight     =   2505
   ScaleWidth      =   2460
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   1560
      Width           =   975
   End
End
Attribute VB_Name = "AnalocClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub DrawLine(lHDC As Long, X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer, lColor As Long)
    Me.Line (X1 * 15, Y1 * 15)-(X2 * 15, Y2 * 15), lColor
    Me.Refresh
End Sub

Public Sub DrawClock()
Dim Angle As Double, X As Integer, Y As Integer, I As Integer
Dim TwoPI As Double, Sec As Byte, Min As Byte, Hr As Byte

Sec = Second(Time)
Min = Minute(Time)
Hr = Hour(Time)
TwoPI = (2 * (4 * Atn(1)))

' Clear the view
Cls

' Draw the seconds arm
    Angle = Sec * TwoPI / 60
    Y = (85 - 70 * Cos(Angle))
    X = (85 + 70 * Sin(Angle))
    DrawLine Me.hDC, 85, 85, X, Y, vbBlack

' Draw the minutes arm
          Angle = Min * TwoPI / 60
          Y = (85 - 65 * Cos(Angle))
          X = (85 + 65 * Sin(Angle))
          DrawLine Me.hDC, 85, 85 - 2, X, Y, vbBlue
          DrawLine Me.hDC, 85, 85, X, Y, vbBlue
          DrawLine Me.hDC, 85, 85 + 2, X, Y, vbBlue

' Draw the hours arm
          Angle = ((Hr Mod 12) * 60 - 2 + Min) * TwoPI / 12 / 60
          Y = (85 - 50 * Cos(Angle))
          X = (85 + 50 * Sin(Angle))
          DrawLine Me.hDC, 85, 85 - 2, X, Y, vbRed
          DrawLine Me.hDC, 85, 85, X, Y, vbRed
          DrawLine Me.hDC, 85, 85 + 2, X, Y, vbRed
End Sub

Private Sub Form_Load()
   DrawClock
End Sub

Private Sub Timer1_Timer()
DrawClock
Label1.Caption = Time
End Sub
