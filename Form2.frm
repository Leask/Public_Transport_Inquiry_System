VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13845
   LinkTopic       =   "Form2"
   ScaleHeight     =   564
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   923
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   825
      Left            =   10935
      TabIndex        =   1
      Top             =   7245
      Width           =   2490
   End
   Begin VB.PictureBox Map_Frm 
      AutoRedraw      =   -1  'True
      Height          =   6630
      Left            =   315
      ScaleHeight     =   438
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   863
      TabIndex        =   0
      Top             =   180
      Width           =   13005
      Begin VB.Image Map_Image 
         Appearance      =   0  'Flat
         Height          =   4500
         Index           =   11
         Left            =   10575
         Stretch         =   -1  'True
         Top             =   6435
         Width           =   4500
      End
      Begin VB.Image Map_Image 
         Appearance      =   0  'Flat
         Height          =   4500
         Index           =   10
         Left            =   9000
         Stretch         =   -1  'True
         Top             =   4455
         Width           =   4500
      End
      Begin VB.Image Map_Image 
         Appearance      =   0  'Flat
         Height          =   4500
         Index           =   9
         Left            =   0
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   4500
      End
      Begin VB.Image Map_Image 
         Appearance      =   0  'Flat
         Height          =   4500
         Index           =   8
         Left            =   225
         Stretch         =   -1  'True
         Top             =   5490
         Width           =   4500
      End
      Begin VB.Image Map_Image 
         Appearance      =   0  'Flat
         Height          =   4500
         Index           =   7
         Left            =   10575
         Stretch         =   -1  'True
         Top             =   495
         Width           =   4500
      End
      Begin VB.Image Map_Image 
         Appearance      =   0  'Flat
         Height          =   4500
         Index           =   6
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4500
      End
      Begin VB.Image Map_Image 
         Appearance      =   0  'Flat
         Height          =   4500
         Index           =   5
         Left            =   4500
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4500
      End
      Begin VB.Image Map_Image 
         Appearance      =   0  'Flat
         Height          =   4500
         Index           =   4
         Left            =   10440
         Stretch         =   -1  'True
         Top             =   3375
         Width           =   4500
      End
      Begin VB.Image Map_Image 
         Appearance      =   0  'Flat
         Height          =   4500
         Index           =   3
         Left            =   8280
         Stretch         =   -1  'True
         Top             =   315
         Width           =   4500
      End
      Begin VB.Image Map_Image 
         Appearance      =   0  'Flat
         Height          =   4500
         Index           =   2
         Left            =   4500
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   4500
      End
      Begin VB.Image Map_Image 
         Appearance      =   0  'Flat
         Height          =   4500
         Index           =   1
         Left            =   7920
         Stretch         =   -1  'True
         Top             =   6525
         Width           =   4500
      End
      Begin VB.Image Map_Image 
         Appearance      =   0  'Flat
         Height          =   4500
         Index           =   0
         Left            =   4995
         Stretch         =   -1  'True
         Top             =   6075
         Width           =   4500
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================
'地图显示模块 DEMO 2
'
'设计/编程: 黄思夏
'Design/Program: Leask Huang
'Copyfree (C) 2008 Leask Huang
'
'==============================================================


Option Explicit   '强制声明变量
Dim Map_Left As Integer   '地图切片左界限
Dim Map_Right As Integer   '地图切片右界限
Dim Map_Up As Integer   '地图切片上界限
Dim Map_Down As Integer   '地图切片下界限
Dim Cur_City_Name As String   '当前城市
Dim Map_Cou_Width As Long
Dim Map_Cou_Height As Long
Dim Map_Bas_X As Long
Dim Map_Bas_Y As Long
Dim Map_Cur_X As Long
Dim Map_Cur_Y As Long
Dim Map_Max_W As Long
Dim Map_Max_H As Long
Dim Map_Move_X As Integer
Dim Map_Move_Y As Integer
Dim Map_Square(11, 2) As Integer



Private Sub Map_Exp(X As Long, Y As Long)
Dim Show_X As Long
Dim Show_Y As Long
Dim Cell_X As Long
Dim Cell_Y As Long
Dim ix As Integer
Dim iy As Integer
Dim iDx As Integer
Dim iDy As Integer
Dim i As Integer

If X < 0 Or X > Map_Max_W Or Y < 0 Or Y > Map_Max_H Then Exit Sub

Map_Image(0).Top = 0
Map_Image(0).Left = 0

Map_Image(1).Top = 0
Map_Image(1).Left = 300

Map_Image(2).Top = 0
Map_Image(2).Left = 600

Map_Image(3).Top = 0
Map_Image(3).Left = 900

Map_Image(4).Top = 300
Map_Image(4).Left = 0

Map_Image(5).Top = 300
Map_Image(5).Left = 300

Map_Image(6).Top = 300
Map_Image(6).Left = 600

Map_Image(7).Top = 300
Map_Image(7).Left = 900

Map_Image(8).Top = 600
Map_Image(8).Left = 0

Map_Image(9).Top = 600
Map_Image(9).Left = 300

Map_Image(10).Top = 600
Map_Image(10).Left = 600

Map_Image(11).Top = 600
Map_Image(11).Left = 900

Show_X = X
Show_Y = Y

Cell_X = Show_X \ 300
Cell_Y = Show_Y \ 300

If Show_X - Cell_X * 300 < 150 Then
iDx = -2
Else
iDx = -1
End If

If Show_Y - Cell_Y * 300 < 150 Then
iDy = -1
Else
iDy = 0
End If

For iy = 0 To 2
    For ix = 0 To 3
        Load_Map_Image Cell_X + ix - iDx, Cell_Y + iy - iDy, i
        i = i + 1
    Next
Next

ix = (Map_Frm.Width - 1200) / 2 - (Show_X - Cell_X * 300)
iy = (Map_Frm.Height - 900) / 2 - (Show_Y - Cell_Y * 300)

For i = 0 To 11
Map_Image(i).Left = Map_Image(i).Left + ix
Map_Image(i).Top = Map_Image(i).Top + iy
Next


End Sub

Private Sub Map_Init()
Dim i As Integer
For i = 0 To 11
Map_Image(i).Stretch = False
Map_Image(i).Width = 300
Map_Image(i).Height = 300
Map_Square(i, 0) = i
Next

Map_Bas_X = 22753
Map_Bas_Y = 39284

Map_Cou_Width = 171 - 1
Map_Cou_Height = 106 - 1

Map_Max_W = Map_Cou_Width * 300
Map_Max_H = Map_Cou_Height * 300



Cur_City_Name = "Shenzhen"
End Sub

Private Function Get_Map_Path()   '获得当前城市的地图数据路径
Get_Map_Path = App.Path & "\Library\Maps\" & Cur_City_Name & "\"
End Function

Private Function Get_Map_Image(X As Integer, Y As Integer)   '获得地图切片路径
Dim temp_Path As String
temp_Path = Get_Map_Path & (Y + Map_Bas_Y) & "\" & (Y + Map_Bas_Y) & "_" & (X + Map_Bas_X) & ".bmp"
Dim FSO As New FileSystemObject
If FSO.FileExists(temp_Path) = False Then temp_Path = App.Path & "\Library\Images\space.bmp"
Get_Map_Image = temp_Path
End Function

Private Sub Load_Map_Image(X As Integer, Y As Integer, Map_Square_ID As Integer)   '加载地图切片
Map_Image(Map_Square_ID).Picture = LoadPicture(Get_Map_Image(X, Y))
Map_Square(Map_Square_ID, 1) = X
Map_Square(Map_Square_ID, 2) = Y
End Sub

Private Sub Form_Load()
Map_Init
Map_Exp 9800, 8000
End Sub

Private Sub Map_Image_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Map_Move_X = X
Map_Move_Y = Y
End If
End Sub

Private Sub Map_Image_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Caption = X & " " & Y
Dim i As Integer
If Button = 1 Then
    For i = 0 To 11
        Map_Image(i).Left = Map_Image(i).Left + (X - Map_Move_X) / 15
        Map_Image(i).Top = Map_Image(i).Top + (Y - Map_Move_Y) / 15
    Next
    
If Map_Image(Map_Square(0, 0)).Left > 0 Then
    Map_Square(0, 0) = Map_Square(3, 0)
    Map_Square(1, 0) = Map_Square(0, 0)
    Map_Square(2, 0) = Map_Square(0, 0)
    Map_Square(3, 0) = Map_Square(0, 0)
    Map_Square(4, 0) = Map_Square(0, 0)
    Map_Square(5, 0) = Map_Square(0, 0)
    Map_Square(6, 0) = Map_Square(0, 0)
    Map_Square(7, 0) = Map_Square(0, 0)
    Map_Square(8, 0) = Map_Square(0, 0)
    Map_Square(9, 0) = Map_Square(0, 0)
    Map_Square(10, 0) = Map_Square(0, 0)
    Map_Square(11, 0) = Map_Square(0, 0)
End If
    
If Map_Image(Map_Square(0, 0)).Top > 0 Then
End If

If Map_Image(Map_Square(11, 0)).Left + 300 < Map_Frm.Width Then
End If

If Map_Image(Map_Square(11, 0)).Top + 300 < Map_Frm.Height Then
End If
    
    
    
End If
End Sub
