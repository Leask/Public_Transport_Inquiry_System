VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14325
   LinkTopic       =   "Form1"
   ScaleHeight     =   8940
   ScaleWidth      =   14325
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.HScrollBar Map_CT_X 
      Height          =   280
      LargeChange     =   1000
      Left            =   90
      SmallChange     =   100
      TabIndex        =   5
      Top             =   6795
      Width           =   10185
   End
   Begin VB.VScrollBar Map_CT_Y 
      Height          =   6675
      LargeChange     =   1000
      Left            =   10305
      SmallChange     =   100
      TabIndex        =   4
      Top             =   90
      Width           =   280
   End
   Begin VB.PictureBox Map_FM 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6675
      Left            =   90
      ScaleHeight     =   6645
      ScaleWidth      =   10155
      TabIndex        =   2
      Top             =   90
      Width           =   10185
      Begin VB.PictureBox Map_CT 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   20535
         Left            =   2205
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   20535
         ScaleWidth      =   30720
         TabIndex        =   3
         Top             =   1710
         Width           =   30720
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   240
      Index           =   1
      Left            =   11070
      TabIndex        =   1
      Top             =   720
      Width           =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   240
      Index           =   0
      Left            =   11070
      TabIndex        =   0
      Top             =   270
      Width           =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Demo 1
'=====================Leask==========================
'(C) 2007 Leask Huang
'
'

Option Explicit
Dim Move_X As Long
Dim Move_Y As Long

Private Sub Map_Init()
Map_CT_X.Top = Map_FM.Top + Map_FM.Height + 10
Map_CT_Y.Left = Map_FM.Left + Map_FM.Width + 10
Map_CT_X.Height = 280
Map_CT_Y.Width = 280
Map_CT_X.Width = Map_FM.Width
Map_CT_Y.Height = Map_FM.Height
Map_CT_X.Left = Map_FM.Left
Map_CT_Y.Top = Map_FM.Top
Map_CT_X.Min = 0
Map_CT_X.Max = Map_FM.Width - Map_CT.Width
Map_CT_Y.Min = 0
Map_CT_Y.Max = Map_FM.Height - Map_CT.Height
Map_CT.Left = 0
Map_CT.Top = 0
End Sub

Private Sub Form_Load()
Map_Init
End Sub

Private Sub Map_CT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Move_X = X
Move_Y = Y
End If
End Sub

Private Sub Map_CT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Dim Move_X_Do As Long
Dim Move_Y_Do As Long
Move_X_Do = Map_CT.Left + X - Move_X
Move_Y_Do = Map_CT.Top + Y - Move_Y
If Move_X_Do > 0 Then Move_X_Do = 0
If Move_X_Do + Map_CT.Width - Map_FM.Width < 0 Then Move_X_Do = Map_FM.Width - Map_CT.Width
If Move_Y_Do + Map_CT.Height - Map_FM.Height < 0 Then Move_Y_Do = Map_FM.Height - Map_CT.Height
If Move_Y_Do > 0 Then Move_Y_Do = 0
Map_CT.Left = Move_X_Do
Map_CT.Top = Move_Y_Do
Map_CT_X.Value = Map_CT.Left
Map_CT_Y.Value = Map_CT.Top
End If
Label1(0).Caption = "X: " & X
Label1(1).Caption = "Y: " & Y
End Sub
Private Sub Map_CT_x_Scroll()
Map_CT.Left = Map_CT_X.Value
End Sub

Private Sub Map_CT_Y_Scroll()
Map_CT.Top = Map_CT_Y.Value
End Sub
