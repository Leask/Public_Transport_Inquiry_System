VERSION 5.00
Begin VB.Form Form_Main 
   Caption         =   "智能公交查询系统"
   ClientHeight    =   9090
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14730
   Icon            =   "Form_Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   606
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   982
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10530
      Top             =   8460
   End
   Begin VB.TextBox Text_RS 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   135
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   44
      Top             =   5625
      Width           =   5865
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4785
      Left            =   135
      TabIndex        =   23
      Top             =   630
      Visible         =   0   'False
      Width           =   5910
      Begin VB.ListBox List2 
         Enabled         =   0   'False
         Height          =   1140
         Left            =   2610
         TabIndex        =   29
         Top             =   540
         Width           =   3210
      End
      Begin VB.TextBox Text1 
         Height          =   2670
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   2025
         Width           =   5685
      End
      Begin VB.CommandButton CMB 
         Caption         =   "下行"
         Height          =   420
         Index           =   1
         Left            =   90
         TabIndex        =   26
         Top             =   1035
         Width           =   645
      End
      Begin VB.CommandButton CMB 
         Caption         =   "上行"
         Height          =   420
         Index           =   0
         Left            =   90
         TabIndex        =   25
         Top             =   540
         Width           =   645
      End
      Begin VB.ListBox List1 
         Enabled         =   0   'False
         Height          =   1140
         Left            =   765
         TabIndex        =   24
         Top             =   540
         Width           =   1770
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   195
         Left            =   90
         TabIndex        =   27
         Top             =   1755
         Width           =   1185
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "线路查询、站点查询、地点查询"
      Height          =   330
      Index           =   1
      Left            =   2745
      TabIndex        =   15
      Top             =   135
      Width           =   3345
   End
   Begin VB.CommandButton Command2 
      Caption         =   "公交换乘指南"
      Height          =   330
      Index           =   0
      Left            =   135
      TabIndex        =   14
      Top             =   135
      Width           =   2400
   End
   Begin VB.Frame Frame_CT 
      Caption         =   "线路查询、地点查询、站点查询"
      Height          =   4785
      Index           =   1
      Left            =   135
      TabIndex        =   11
      Top             =   630
      Visible         =   0   'False
      Width           =   5910
      Begin VB.ListBox List3 
         Height          =   1500
         ItemData        =   "Form_Main.frx":57E2
         Left            =   3195
         List            =   "Form_Main.frx":57E4
         TabIndex        =   36
         Top             =   3015
         Width           =   2490
      End
      Begin VB.CommandButton Command1 
         Caption         =   "<<"
         Height          =   285
         Index           =   7
         Left            =   2475
         TabIndex        =   22
         Top             =   2610
         Width           =   330
      End
      Begin VB.CommandButton Command1 
         Caption         =   "<<"
         Height          =   285
         Index           =   4
         Left            =   5355
         TabIndex        =   21
         Top             =   315
         Width           =   330
      End
      Begin VB.ListBox List_Input 
         Height          =   1140
         Index           =   4
         Left            =   3195
         TabIndex        =   20
         Top             =   1080
         Width           =   2490
      End
      Begin VB.TextBox Text_Input 
         Height          =   285
         Index           =   4
         Left            =   3825
         TabIndex        =   19
         Top             =   315
         Width           =   1400
      End
      Begin VB.TextBox Text_Input 
         Height          =   285
         Index           =   3
         Left            =   900
         TabIndex        =   18
         Top             =   2610
         Width           =   1400
      End
      Begin VB.ListBox List_Input 
         Height          =   1140
         Index           =   3
         Left            =   270
         TabIndex        =   17
         Top             =   3375
         Width           =   2495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "<<"
         Height          =   285
         Index           =   2
         Left            =   2385
         TabIndex        =   16
         Top             =   315
         Width           =   330
      End
      Begin VB.TextBox Text_Input 
         Height          =   285
         Index           =   2
         Left            =   855
         TabIndex        =   13
         Top             =   315
         Width           =   1400
      End
      Begin VB.ListBox List_Input 
         Height          =   1140
         Index           =   2
         Left            =   225
         TabIndex        =   12
         Top             =   1080
         Width           =   2490
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "参考站点："
         Height          =   180
         Index           =   6
         Left            =   270
         TabIndex        =   42
         Top             =   3060
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "站点：                           参考路线："
         Height          =   180
         Index           =   5
         Left            =   270
         TabIndex        =   41
         Top             =   2655
         Width           =   3870
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   5715
         X2              =   225
         Y1              =   2430
         Y2              =   2430
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   2925
         X2              =   2925
         Y1              =   315
         Y2              =   2205
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "参考线路：                       参考地点："
         Height          =   180
         Index           =   4
         Left            =   225
         TabIndex        =   35
         Top             =   765
         Width           =   3870
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "线路：                           地点："
         Height          =   180
         Index           =   3
         Left            =   225
         TabIndex        =   34
         Top             =   360
         Width           =   3510
      End
   End
   Begin VB.Frame Frame_CT 
      Caption         =   "公交换乘指南"
      Height          =   4785
      Index           =   0
      Left            =   135
      TabIndex        =   4
      Top             =   630
      Width           =   5910
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "清除(&C)"
         Height          =   330
         Index           =   5
         Left            =   1575
         TabIndex        =   48
         Top             =   2385
         Width           =   1140
      End
      Begin VB.CommandButton Command1 
         Caption         =   "<<"
         Height          =   285
         Index           =   3
         Left            =   5355
         TabIndex        =   43
         Top             =   315
         Width           =   330
      End
      Begin VB.ListBox List_Output 
         Height          =   1680
         ItemData        =   "Form_Main.frx":57E6
         Left            =   225
         List            =   "Form_Main.frx":57E8
         TabIndex        =   32
         Top             =   2880
         Width           =   5460
      End
      Begin VB.CommandButton Command1 
         Caption         =   "<<"
         Height          =   285
         Index           =   1
         Left            =   2385
         TabIndex        =   10
         Top             =   315
         Width           =   330
      End
      Begin VB.CommandButton Command1 
         Caption         =   "查询(&Y)"
         Default         =   -1  'True
         Height          =   330
         Index           =   0
         Left            =   225
         TabIndex        =   9
         Top             =   2385
         Width           =   1140
      End
      Begin VB.ListBox List_Input 
         Height          =   1140
         Index           =   1
         Left            =   3195
         TabIndex        =   8
         Top             =   1080
         Width           =   2490
      End
      Begin VB.TextBox Text_Input 
         Height          =   285
         Index           =   1
         Left            =   3825
         TabIndex        =   7
         Top             =   315
         Width           =   1400
      End
      Begin VB.TextBox Text_Input 
         Height          =   285
         Index           =   0
         Left            =   855
         TabIndex        =   6
         Top             =   315
         Width           =   1400
      End
      Begin VB.ListBox List_Input 
         Height          =   1140
         Index           =   0
         Left            =   225
         TabIndex        =   5
         Top             =   1080
         Width           =   2490
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "参考路线："
         Height          =   180
         Index           =   2
         Left            =   2970
         TabIndex        =   33
         Top             =   2475
         Width           =   900
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   2925
         X2              =   2925
         Y1              =   315
         Y2              =   2205
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "参考站点：                       参考站点："
         Height          =   180
         Index           =   1
         Left            =   225
         TabIndex        =   31
         Top             =   765
         Width           =   3870
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "起点：                           终点："
         Height          =   180
         Index           =   0
         Left            =   225
         TabIndex        =   30
         Top             =   360
         Width           =   3510
      End
   End
   Begin VB.HScrollBar Map_CT_X 
      Height          =   280
      LargeChange     =   1000
      Left            =   11655
      SmallChange     =   100
      TabIndex        =   3
      Top             =   8010
      Width           =   7305
   End
   Begin VB.VScrollBar Map_CT_Y 
      Height          =   5145
      LargeChange     =   1000
      Left            =   17685
      SmallChange     =   100
      TabIndex        =   2
      Top             =   1215
      Width           =   280
   End
   Begin VB.PictureBox Map_FM 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7800
      Left            =   6210
      ScaleHeight     =   518
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   689
      TabIndex        =   0
      Top             =   135
      Width           =   10365
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4905
         Left            =   3375
         ScaleHeight     =   4905
         ScaleWidth      =   6090
         TabIndex        =   49
         Top             =   1170
         Visible         =   0   'False
         Width           =   6090
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4815
            Left            =   45
            ScaleHeight     =   4815
            ScaleWidth      =   6000
            TabIndex        =   50
            Top             =   45
            Width           =   6000
            Begin VB.Label Label8 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   11.25
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   540
               TabIndex        =   51
               Top             =   4549
               Width           =   120
            End
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   90
         ScaleHeight     =   615
         ScaleWidth      =   1965
         TabIndex        =   45
         Top             =   90
         Width           =   1995
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   90
            TabIndex        =   47
            Top             =   360
            Width           =   90
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "广东省 - 深圳市"
            Height          =   180
            Left            =   90
            TabIndex        =   46
            Top             =   90
            Width           =   1350
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   8460
         ScaleHeight     =   615
         ScaleWidth      =   1740
         TabIndex        =   37
         Top             =   90
         Visible         =   0   'False
         Width           =   1770
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Caption         =   "下行"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   1
            Left            =   990
            TabIndex        =   39
            Top             =   315
            Width           =   735
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Caption         =   "上行"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   0
            Left            =   90
            TabIndex        =   38
            Top             =   315
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "路线：329"
            Height          =   180
            Left            =   90
            TabIndex        =   40
            Top             =   90
            Width           =   810
         End
      End
      Begin VB.PictureBox Map_CT 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   54000
         Left            =   -135
         MousePointer    =   5  'Size
         ScaleHeight     =   3600
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   5400
         TabIndex        =   1
         Top             =   1395
         Width           =   81000
         Begin VB.Line Line2 
            BorderWidth     =   3
            X1              =   15
            X2              =   75
            Y1              =   414
            Y2              =   414
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1 公里"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   360
            TabIndex        =   53
            Top             =   5940
            Width           =   600
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "麦当劳"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   2340
            TabIndex        =   52
            Top             =   5175
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Shape Shape_CS 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H000000FF&
            BorderWidth     =   2
            FillColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   0
            Left            =   2520
            Shape           =   3  'Circle
            Top             =   2655
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.Line Line_Cell 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   5
            Index           =   0
            Visible         =   0   'False
            X1              =   213
            X2              =   423
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Shape Shape_LC 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   5
            Height          =   1500
            Left            =   1395
            Shape           =   3  'Circle
            Top             =   810
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H0080FFFF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FF0000&
            BorderStyle     =   3  'Dot
            Height          =   375
            Left            =   2205
            Top             =   5085
            Visible         =   0   'False
            Width           =   780
         End
      End
   End
End
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================
'智能公交查询系统 Demo7
'
'设计/编程: 黄思夏(Leask Huang)
'数据库: 梁健洲(Len Leung)
'
'Copyfree (C) 2008 Leask Huang
'www.leaskh.com
'leaskh@gmail.com
'
'==============================================================


Option Explicit
    Dim Move_X As Long
    Dim Move_Y As Long
    Dim TX_Ch As Boolean
    Private Rs As New ADODB.Recordset
    Private Conn As New ADODB.Connection
    Dim Result_Arr() As String
    Dim Result_Arr_i As Integer
    Dim Draw_Str(2) As String
    Dim Count_Close As Integer
    Dim Path_Mode As Integer
    Dim Temp_X As Long
    Dim Temp_Y As Long
    Dim Load_CI As Integer
    Dim Load_LCAB As Integer


Private Sub Map_Init()
    Map_CT.Picture = LoadPicture(App.Path & "\Library\Shenzhen\city_map.gif")
    Map_CT.Left = (Map_FM.Width - Map_CT.Width) / 2
    Map_CT.Top = (Map_FM.Height - Map_CT.Height) / 2
    Map_CT_X.Top = Map_FM.Top + Map_FM.Height + 2
    Map_CT_Y.Left = Map_FM.Left + Map_FM.Width + 2
    Map_CT_X.Height = 17
    Map_CT_Y.Width = 17
    Map_CT_X.Width = Map_FM.Width
    Map_CT_Y.Height = Map_FM.Height
    Map_CT_X.Left = Map_FM.Left
    Map_CT_Y.Top = Map_FM.Top
    Map_CT_X.Min = 0
    Map_CT_X.Max = Map_FM.Width - Map_CT.Width
    Map_CT_Y.Min = 0
    Map_CT_Y.Max = Map_FM.Height - Map_CT.Height
    Map_CT_X.Value = Map_CT.Left
    Map_CT_Y.Value = Map_CT.Top
End Sub


Private Sub CMB_Click(Index As Integer)
Dim Arr_In() As String
Dim i As Integer
    Select Case Index
        Case 0
            RS_Init
            Rs.Open "SELECT * FROM Line_Info WHERE Line_Num LIKE '" & List_Input(2).List(List_Input(2).ListIndex) & "'", Conn
            Arr_In() = Split(Rs.Fields.Item(7), ">")
            Path_Mode = 1
            Label3.Caption = "上行"
        Case 1
            RS_Init
            Rs.Open "SELECT * FROM Line_Info WHERE Line_Num LIKE '" & List_Input(2).List(List_Input(2).ListIndex) & "'", Conn
            Arr_In() = Split(Rs.Fields.Item(8), ">")
            Path_Mode = 2
            Label3.Caption = "下行"
    End Select
    
    List1.Clear
    List2.Clear
    
    For i = 0 To UBound(Arr_In)
        List1.AddItem Arr_In(i)
        List2.AddItem " "
    Next
    
    List1.ListIndex = 0
    List2.ListIndex = 0
    
    Temp_X = 0
    Temp_Y = 0
    
End Sub

Private Sub Command1_Click(Index As Integer)
On Error Resume Next
Dim Do_i As Integer
Dim Check_Ed As Boolean
Dim Do_ii As Integer
Dim i As Integer
Dim ii As Integer
Dim iii As Integer
Dim Line_Colt() As String
Dim Line_Colt_A() As String
Dim Str_SPL() As String
Dim Str_SPL_A() As String
Dim Str_SPL_B() As String
Dim Str_SPL_C() As String
Dim Str_SPL_D() As String
Dim Str_i As Integer
Dim ST_OK As Boolean
Dim Find_Si As Integer
Dim Find_OK As Boolean
   
    Select Case Index
        Case 0
            Result_Arr_i = 0
            ReDim Result_Arr(99, 9) As String
            Check_Ed = False
            RS_Init
            Rs.Open "SELECT * FROM Temp_Shed WHERE (St_A LIKE '" & Text_Input(0).Text & "') AND (St_B LIKE '" & Text_Input(1).Text & "')", Conn
            Do Until Rs.EOF = True
                Check_Ed = True
                Rs.MoveNext
            Loop
            
            If Check_Ed = False Then
                
                For Do_i = 0 To 1
                    RS_Init
                    Select Case Do_i
                        Case 0: Rs.Open "SELECT * FROM Line_Info WHERE Stations_Up_Going LIKE '%" & Text_Input(0).Text & "%" & Text_Input(1).Text & "%'", Conn
                        Case 1: Rs.Open "SELECT * FROM Line_Info WHERE Stations_Down_Going LIKE '%" & Text_Input(0).Text & "%" & Text_Input(1).Text & "%'", Conn
                    End Select
                    Do Until Rs.EOF = True
                        Result_Arr(Result_Arr_i, 0) = Text_Input(0).Text
                        Result_Arr(Result_Arr_i, 1) = Text_Input(1).Text
                        Result_Arr(Result_Arr_i, 2) = False
                        Result_Arr(Result_Arr_i, 3) = Rs.Fields.Item(1)
                        Select Case Do_i
                            Case 0
                                Str_SPL = Split(Rs.Fields.Item(7), ">")
                                Str_SPL_C = Split(Rs.Fields.Item(10), ">")
                            Case 1
                                Str_SPL = Split(Rs.Fields.Item(8), ">")
                                Str_SPL_C = Split(Rs.Fields.Item(11), ">")
                        End Select
                        ST_OK = False
                        For i = 0 To UBound(Str_SPL)
                            If InStr(Str_SPL(i), Text_Input(0).Text) > 0 Then ST_OK = True
                            If InStr(Str_SPL(i), Text_Input(1).Text) > 0 And ST_OK = True Then
                                Result_Arr(Result_Arr_i, 5) = Result_Arr(Result_Arr_i, 5) & Str_SPL(i)
                                Result_Arr(Result_Arr_i, 7) = Result_Arr(Result_Arr_i, 7) & Str_SPL_C(i)
                                Exit For
                            End If
                            If ST_OK = True Then
                                Result_Arr(Result_Arr_i, 5) = Result_Arr(Result_Arr_i, 5) & Str_SPL(i) & ">"
                                Result_Arr(Result_Arr_i, 7) = Result_Arr(Result_Arr_i, 7) & Str_SPL_C(i) & ">"
                            End If
                        Next
                        Conn.Execute "INSERT INTO Temp_Shed (St_A,St_B,Line_CH,Line_A,Line_A_Sts,Line_A_Path,Line_Long) VALUES ('" & Result_Arr(Result_Arr_i, 0) & "','" & Result_Arr(Result_Arr_i, 1) & "'," & Result_Arr(Result_Arr_i, 2) & ",'" & Result_Arr(Result_Arr_i, 3) & "','" & Result_Arr(Result_Arr_i, 5) & "','" & Result_Arr(Result_Arr_i, 7) & "'," & Count_Path(Result_Arr(Result_Arr_i, 7), "", False) & ")"
                        Result_Arr_i = Result_Arr_i + 1
                        Rs.MoveNext
                    Loop
                Next
                
                For Do_i = 0 To 1
                    RS_Init
                    Select Case Do_i
                        Case 0: Rs.Open "SELECT * FROM Line_Info WHERE (Stations_Up_Going LIKE '%" & Text_Input(0).Text & "%') and (Not Stations_Up_Going LIKE '%" & Text_Input(1).Text & "%')", Conn
                        Case 1: Rs.Open "SELECT * FROM Line_Info WHERE (Stations_Down_Going LIKE '%" & Text_Input(0).Text & "%') and (Not Stations_Down_Going LIKE '%" & Text_Input(1).Text & "%')", Conn
                    End Select
                    Str_i = 0
                    ReDim Line_Colt(99, 1) As String
                    ReDim Line_Colt_A(99, 1) As String
                    Do Until Rs.EOF = True
                        Line_Colt(Str_i, 0) = Rs.Fields.Item(1)
                        Select Case Do_i
                            Case 0
                                Line_Colt(Str_i, 1) = Rs.Fields.Item(7)
                                Line_Colt_A(Str_i, 1) = Rs.Fields.Item(10)
                            Case 1
                                Line_Colt(Str_i, 1) = Rs.Fields.Item(8)
                                Line_Colt_A(Str_i, 1) = Rs.Fields.Item(11)
                        End Select
                        Str_i = Str_i + 1
                        Rs.MoveNext
                    Loop
                    For i = 0 To 99
                        If Line_Colt(i, 0) = "" Then Exit For
                        Str_SPL() = Split(Line_Colt(i, 1), ">")
                        Str_SPL_C() = Split(Line_Colt_A(i, 1), ">")
                        Find_Si = -1
                        Find_OK = False
                        For ii = 0 To UBound(Str_SPL)
                            If Find_OK = False Then
                                If InStr(Str_SPL(ii), Text_Input(0).Text) > 0 Then Find_Si = ii
                                If Find_Si <> -1 And ii > Find_Si Then
                                    For Do_ii = 0 To 1
                                        RS_Init
                                        Select Case Do_ii
                                            Case 0: Rs.Open "SELECT * FROM Line_Info WHERE (Stations_Up_Going LIKE '%" & Str_SPL(ii) & "%" & Text_Input(1).Text & "%') and (Not Stations_Up_Going LIKE '%" & Text_Input(0).Text & "%')", Conn
                                            Case 1: Rs.Open "SELECT * FROM Line_Info WHERE (Stations_Down_Going LIKE '%" & Str_SPL(ii) & "%" & Text_Input(1).Text & "%') and (Not Stations_Down_Going LIKE '%" & Text_Input(0).Text & "%')", Conn
                                        End Select
                                        Do Until Rs.EOF = True
                                            If Result_Arr_i < 99 Then
                                                Find_OK = True
                                                Result_Arr(Result_Arr_i, 0) = Text_Input(0).Text
                                                Result_Arr(Result_Arr_i, 1) = Text_Input(1).Text
                                                Result_Arr(Result_Arr_i, 2) = True
                                                Result_Arr(Result_Arr_i, 3) = Line_Colt(i, 0)
                                                Result_Arr(Result_Arr_i, 4) = Rs.Fields.Item(1)
                                                For iii = Find_Si To ii - 1
                                                    Result_Arr(Result_Arr_i, 5) = Result_Arr(Result_Arr_i, 5) & Str_SPL(iii) & ">"
                                                    Result_Arr(Result_Arr_i, 7) = Result_Arr(Result_Arr_i, 7) & Str_SPL_C(iii) & ">"
                                                Next
                                                Result_Arr(Result_Arr_i, 5) = Result_Arr(Result_Arr_i, 5) & Str_SPL(ii)
                                                Result_Arr(Result_Arr_i, 7) = Result_Arr(Result_Arr_i, 7) & Str_SPL_C(ii)
                                                Select Case Do_ii
                                                    Case 0
                                                        Str_SPL_A = Split(Rs.Fields.Item(7), ">")
                                                        Str_SPL_D = Split(Rs.Fields.Item(10), ">")
                                                    Case 1
                                                        Str_SPL_A = Split(Rs.Fields.Item(8), ">")
                                                        Str_SPL_D = Split(Rs.Fields.Item(11), ">")
                                                End Select
                                                ST_OK = False
                                                For iii = 0 To UBound(Str_SPL_A)
                                                    If InStr(Str_SPL_A(iii), Str_SPL(ii)) > 0 Then ST_OK = True
                                                    If InStr(Str_SPL_A(iii), Text_Input(1).Text) > 0 And ST_OK = True Then
                                                        Result_Arr(Result_Arr_i, 6) = Result_Arr(Result_Arr_i, 6) & Str_SPL_A(iii)
                                                        Result_Arr(Result_Arr_i, 8) = Result_Arr(Result_Arr_i, 8) & Str_SPL_D(iii)
                                                        Exit For
                                                    End If
                                                    If ST_OK = True Then
                                                        Result_Arr(Result_Arr_i, 6) = Result_Arr(Result_Arr_i, 6) & Str_SPL_A(iii) & ">"
                                                        Result_Arr(Result_Arr_i, 8) = Result_Arr(Result_Arr_i, 8) & Str_SPL_D(iii) & ">"
                                                    End If
                                                Next
                                                Conn.Execute "INSERT INTO Temp_Shed (St_A,St_B,Line_CH,Line_A,Line_B,Line_A_Sts,Line_B_Sts,Line_A_Path,Line_B_Path,Line_Long) VALUES ('" & Result_Arr(Result_Arr_i, 0) & "','" & Result_Arr(Result_Arr_i, 1) & "'," & Result_Arr(Result_Arr_i, 2) & ",'" & Result_Arr(Result_Arr_i, 3) & "','" & Result_Arr(Result_Arr_i, 4) & "','" & Result_Arr(Result_Arr_i, 5) & "','" & Result_Arr(Result_Arr_i, 6) & "','" & Result_Arr(Result_Arr_i, 7) & "','" & Result_Arr(Result_Arr_i, 8) & "'," & Count_Path(Result_Arr(Result_Arr_i, 7), Result_Arr(Result_Arr_i, 8), True) & ")"
                                                Result_Arr_i = Result_Arr_i + 1
                                            End If
                                            Rs.MoveNext
                                        Loop
                                    Next
                                End If
                            End If
                        Next
                    Next
                Next
            End If
            
            Result_Arr_i = 0
            
            ReDim Result_Arr(99, 9) As String
            
            RS_Init
            Rs.Open "SELECT * FROM Temp_Shed WHERE (St_A LIKE '" & Text_Input(0).Text & "') AND (St_B LIKE '" & Text_Input(1).Text & "') ORDER BY Line_Long", Conn
            List_Output.Clear
            
            Text_RS.Text = ""
            
            Do Until Rs.EOF = True
                For i = 1 To 10
                    Result_Arr(Result_Arr_i, i - 1) = Rs.Fields.Item(i)
                Next
                Select Case Result_Arr(Result_Arr_i, 2)
                    Case True
                        List_Output.AddItem "乘坐 " & Result_Arr(Result_Arr_i, 3) & " 转 " & Result_Arr(Result_Arr_i, 4) & " 可到达"
                    Case False
                        List_Output.AddItem "乘坐 " & Result_Arr(Result_Arr_i, 3) & " 可直接到达"
                End Select
                Result_Arr_i = Result_Arr_i + 1
                Rs.MoveNext
            Loop
            Result_Arr_i = Result_Arr_i + 1 - 1
            
            Case 1
                Input_Init 0
                Load_Stations 0
            
            Case 2
                Input_Init 2
                Load_Line
                
            Case 3
                Input_Init 1
                Load_Stations 1
            
            Case 5
                Input_Init 0
                Input_Init 1
                Load_Stations 0
                Load_Stations 1
                
            Case 4
                Input_Init 4
                Load_Locations
                            
            Case 7
                Input_Init 3
                Load_Stations 3
    End Select
End Sub

Private Function Count_Path(Path_Str_A As String, Path_Str_B As String, Path_CH_ENB As Boolean) As Single
Dim i As Integer
Dim ii As Integer
Dim D_T_XX As Long
Dim D_T_YY As Long
Dim Path_Arr() As String
Dim Path_Arr_A() As String
Dim Path_Arr_B(9999) As String
Dim Path_Arr_B_i As Integer
    Count_Path = 0
    If Path_CH_ENB = True Then
            If Len(Path_Str_A) > 0 And Len(Path_Str_B) > 0 Then
                    Path_Arr = Split(Path_Str_A & ">" & Path_Str_B, ">")
                Else
                    Count_Path = 99999
                    Exit Function
            End If
        Else
            If Len(Path_Str_A) > 0 Then
                    Path_Arr = Split(Path_Str_A, ">")
                Else
                    Count_Path = 99999
                    Exit Function
            End If
    End If
    For i = 0 To UBound(Path_Arr)
        Path_Arr_A = Split(Path_Arr(i), ":")
        For ii = 0 To UBound(Path_Arr_A)
            Path_Arr_B(Path_Arr_B_i) = Path_Arr_A(ii)
            Path_Arr_B_i = Path_Arr_B_i + 1
        Next
    Next
    Path_Arr_B_i = Path_Arr_B_i - 1
    
    For i = 0 To Path_Arr_B_i
        Path_Arr = Split(Path_Arr_B(i), ",")
        If i > 0 Then
            Count_Path = Count_Path + (Sqr((Path_Arr(0) - D_T_XX) ^ 2 + (Path_Arr(1) - D_T_YY) ^ 2))
        End If
        D_T_XX = Path_Arr(0)
        D_T_YY = Path_Arr(1)
    Next
End Function



Private Sub Command2_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 1
        Frame_CT(i).Visible = False
    Next
    Frame_CT(Index).Visible = True
End Sub






Private Sub Form_Load()
On Error Resume Next
Dim ver_str As String
Dim FSO As New FileSystemObject
    
    If App.PrevInstance = True Then End
    
    If FSO.FileExists(FSO.GetSpecialFolder(SystemFolder) & "\VB6CHS.DLL") = False Then
    FSO.CopyFile App.Path & "\VB6CHS.DLL", FSO.GetSpecialFolder(SystemFolder) & "\VB6CHS.DLL", True
    MsgBox "自动优化已完成,请重新运行 智能公交查询系统!", vbInformation
    End
    End If
    Set FSO = Nothing
    ver_str = "Demo7"
    Me.Caption = "智能公交查询系统 " & App.Major & "." & App.Minor & "-" & ver_str
    
    Map_Init
    DB_Init
    Load_Stations 0
    Load_Stations 1
    Load_Stations 3
    
    Input_Init 9
    Load_Locations
    Load_Line
End Sub

Private Sub Line_Cls()
Dim i As Integer
    Line_Cell(0).Visible = False
    Shape_CS(0).Visible = False
    Picture1.Visible = False
    If Load_CI > 0 Then
        For i = 1 To Load_CI
            Unload Line_Cell(i)
        Next
    End If
    If Load_LCAB > 0 Then
        For i = 1 To Load_LCAB
            Unload Shape_CS(i)
        Next
    End If
    
    Load_CI = -1
    Load_LCAB = -1

End Sub

Private Sub DB_Init()
Dim strConn As String
    strConn = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Library\Shenzhen\DB_shenzhen.mdb;Persist Security Info=False"
    Conn.CursorLocation = adUseClient
    Conn.Open strConn
End Sub

Private Sub RS_Init()
    If Rs.State <> adStateClosed Then Rs.Close
    Rs.CursorType = adOpenKeyset
    Rs.LockType = adLockOptimistic
End Sub

Private Sub Load_Stations(ListID As Integer)
    RS_Init
    Rs.Open "SELECT Station_Name FROM Stations_Info ORDER BY Station_Name", Conn
    
    List_Input(ListID).Clear
    List_Input(ListID).Clear
    List_Input(ListID).Clear
    
    Do Until Rs.EOF = True
        List_Input(ListID).AddItem Rs.Fields.Item(0)
        Rs.MoveNext
    Loop

End Sub



Private Sub Load_Line()
    RS_Init
    Rs.Open "SELECT Line_Num FROM Line_Info ORDER BY Line_Num", Conn

    List_Input(2).Clear
    
    Do Until Rs.EOF = True
        List_Input(2).AddItem Rs.Fields.Item(0)
        Rs.MoveNext
    Loop

End Sub

Private Sub Load_Locations()
    RS_Init
    Rs.Open "SELECT Location FROM Locations_Info ORDER BY Location", Conn
    List_Input(4).Clear
    
    Do Until Rs.EOF = True
        List_Input(4).AddItem Rs.Fields.Item(0)
        Rs.MoveNext
    Loop

End Sub



Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    If Me.Width < 10000 Then Me.Width = 10000
    If Me.Height < 7000 Then Me.Height = 7000
    
    Text_RS.Height = Me.Height / 15 - 417
    Map_FM.Height = Me.Height / 15 - 70
    Map_FM.Width = Me.Width / 15 - 449
    
    Picture5.Left = (Map_FM.Width - Picture5.Width) / 2
    Picture5.Top = (Map_FM.Height - Picture5.Height) / 2
    
    Picture1.Left = Map_FM.Width - 127
    
    Map_CT_X.Top = Map_FM.Top + Map_FM.Height + 2
    Map_CT_Y.Left = Map_FM.Left + Map_FM.Width + 2
    Map_CT_X.Height = 17
    Map_CT_Y.Width = 17
    Map_CT_X.Width = Map_FM.Width
    Map_CT_Y.Height = Map_FM.Height
    Map_CT_X.Left = Map_FM.Left
    Map_CT_Y.Top = Map_FM.Top
    Map_CT_X.Min = 0
    Map_CT_X.Max = Map_FM.Width - Map_CT.Width
    Map_CT_Y.Min = 0
    Map_CT_Y.Max = Map_FM.Height - Map_CT.Height
    
    Line2.X1 = 15 - Map_CT.Left
    Line2.X2 = Line2.X1 + 60
    Line2.Y1 = Map_FM.Height - Map_CT.Top - 15
    Line2.Y2 = Line2.Y1
    Label7.Left = 26 - Map_CT.Left
    Label7.Top = Map_FM.Height - Map_CT.Top - 35
End Sub

Private Sub Find_Line(Str_Find As String)
On Error Resume Next
    RS_Init
    Rs.Open "SELECT * FROM Line_Info WHERE Line_Num LIKE '" & Str_Find & "'", Conn
    Label5.Caption = "路线： " & Rs.Fields.Item(1)
    Text_RS.Text = "=======  " & Rs.Fields.Item(1) & "  =======" & vbCrLf & vbCrLf
    Text_RS.Text = Text_RS.Text & "起点站： " & Rs.Fields.Item(2) & vbCrLf & vbCrLf
    Text_RS.Text = Text_RS.Text & "终点站： " & Rs.Fields.Item(3) & vbCrLf & vbCrLf
    Text_RS.Text = Text_RS.Text & "所属公司： " & Rs.Fields.Item(4) & vbCrLf & vbCrLf
    Text_RS.Text = Text_RS.Text & "上行途径路段： " & vbCrLf & Replace(Rs.Fields.Item(5), ">", " - ") & vbCrLf & vbCrLf
    Text_RS.Text = Text_RS.Text & "下行途径路段： " & vbCrLf & Replace(Rs.Fields.Item(6), ">", " - ") & vbCrLf & vbCrLf
    Text_RS.Text = Text_RS.Text & "上行途径站点： " & vbCrLf & Replace(Rs.Fields.Item(7), ">", " - ") & vbCrLf & vbCrLf
    Text_RS.Text = Text_RS.Text & "下行途径站点： " & vbCrLf & Replace(Rs.Fields.Item(8), ">", " - ") & vbCrLf & vbCrLf
    If Rs.Fields.Item(9) = True Then Text_RS.Text = Text_RS.Text & "路线总长约 " & Int(Change_DD(Count_Path(Rs.Fields.Item(10), "", False)) * 100) / 100 & " 公里"
    If Rs.Fields.Item(9) = True Then
            Draw_Str(1) = Rs.Fields.Item(10)
            Draw_Str(2) = Rs.Fields.Item(11)
            Option1_Click 0
        Else
            Line_Cls
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Conn.Close
End Sub

Private Sub List_Input_Click(Index As Integer)
    Select Case Index
        Case 0, 1
            Text_Input(Index).Text = List_Input(Index).Text
            
        Case 2
            Find_Line List_Input(Index).List(List_Input(Index).ListIndex)
            
        Case 3
            RS_Init
            Rs.Open "SELECT * FROM Line_Info WHERE (Stations_Up_Going LIKE '%" & List_Input(Index).List(List_Input(Index).ListIndex) & "%') or (Stations_Down_Going LIKE '%" & Text_Input(3).Text & "%')  Order By Line_Num", Conn
            List3.Clear
            Do Until Rs.EOF = True
                List3.AddItem Rs.Fields.Item(1)
                Rs.MoveNext
            Loop
    
        Case 4
            Shape_LC.Visible = False
            Label9.Visible = False
            Shape1.Visible = False
            RS_Init
            Rs.Open "SELECT * FROM Locations_Info WHERE Location LIKE '" & List_Input(Index).List(List_Input(Index).ListIndex) & "'", Conn
            Shape_LC.Move Rs.Fields.Item(2) - 50, Rs.Fields.Item(3) - 50
            Show_Loca Rs.Fields.Item(2), Rs.Fields.Item(3)
            Select Case List_Input(Index).List(List_Input(Index).ListIndex)
                Case "麦当劳"
                    Picture5.Visible = True
                    Picture4.Picture = LoadPicture(App.Path & "\business\mcd.jpg")
                    Label8.Caption = Rs.Fields.Item(4)
                    Count_Close = 0
                    Timer1.Enabled = True
                    Label9.Move Rs.Fields.Item(2) - 18, Rs.Fields.Item(3) - 6
                    Shape1.Move Rs.Fields.Item(2) - 26, Rs.Fields.Item(3) - 13
                    Label9.Visible = True
                    Shape1.Visible = True
                    Exit Sub
                Case "星巴克"
                    Picture5.Visible = True
                    Picture4.Picture = LoadPicture(App.Path & "\business\stb.jpg")
                    Label8.Caption = Rs.Fields.Item(4)
                    Count_Close = 0
                    Timer1.Enabled = True
                    Label9.Move Rs.Fields.Item(2) - 18, Rs.Fields.Item(3) - 6
                    Shape1.Move Rs.Fields.Item(2) - 26, Rs.Fields.Item(3) - 13
                    Label9.Visible = True
                    Shape1.Visible = True
                    Exit Sub
                Case "必胜客"
                    Picture5.Visible = True
                    Picture4.Picture = LoadPicture(App.Path & "\business\pzh.jpg")
                    Label8.Caption = Rs.Fields.Item(4)
                    Count_Close = 0
                    Timer1.Enabled = True
                    Label9.Move Rs.Fields.Item(2) - 18, Rs.Fields.Item(3) - 6
                    Shape1.Move Rs.Fields.Item(2) - 26, Rs.Fields.Item(3) - 13
                    Label9.Visible = True
                    Shape1.Visible = True
                    Exit Sub
            End Select
            Shape_LC.Visible = True
        End Select
End Sub

Private Sub List_Input_GotFocus(Index As Integer)
    TX_Ch = False
End Sub


Private Sub List_Output_Click()
Dim i As Integer
Dim Str_Temp_A() As String
Dim Str_Temp_B() As String
    
    If List_Output.ListCount = 0 Then Exit Sub
    Text_RS.Text = ""
    Str_Temp_A() = Split(Result_Arr(List_Output.ListIndex, 5), ">")
    Str_Temp_B() = Split(Result_Arr(List_Output.ListIndex, 6), ">")
    
    Text_RS.Text = "=======  公交换乘指南  =======" & vbCrLf & vbCrLf
    Text_RS.Text = Text_RS.Text & "在 " & Str_Temp_A(0) & " 站乘坐 " & Result_Arr(List_Output.ListIndex, 3) & " 路车；" & vbCrLf & vbCrLf
    Text_RS.Text = Text_RS.Text & "乘坐" & Str(UBound(Str_Temp_A) - 1) & " 个站：" & Replace(Result_Arr(List_Output.ListIndex, 5), ">", " - ") & "；" & vbCrLf & vbCrLf
    
    Select Case Result_Arr(List_Output.ListIndex, 2)
        
        Case True
            Text_RS.Text = Text_RS.Text & "在 " & Str_Temp_B(0) & " 站换乘 " & Result_Arr(List_Output.ListIndex, 4) & " 路车；" & vbCrLf & vbCrLf
            Text_RS.Text = Text_RS.Text & "乘坐" & Str(UBound(Str_Temp_B) - 1) & " 个站：" & Replace(Result_Arr(List_Output.ListIndex, 6), ">", " - ") & "；" & vbCrLf & vbCrLf
            Text_RS.Text = Text_RS.Text & "在 " & Str_Temp_B(UBound(Str_Temp_B)) & " 站下车即可到达。"
        Case False
            Text_RS.Text = Text_RS.Text & "在 " & Str_Temp_A(UBound(Str_Temp_A)) & " 站下车即可到达。"
    End Select
    
    If Result_Arr(List_Output.ListIndex, 9) <> 99999 Then
            Text_RS.Text = Text_RS.Text & vbCrLf & vbCrLf & "全程约 " & Int(Change_DD(Result_Arr(List_Output.ListIndex, 9)) * 100) / 100 & " 公里。"
            Draw_Str(1) = Result_Arr(List_Output.ListIndex, 7)
            Draw_Str(2) = Result_Arr(List_Output.ListIndex, 8)
            Draw_Str(0) = "STATION"
            Show_Line
        Else
            Line_Cls
    End If
End Sub

Function Change_DD(Long_St As String) As Single
    Change_DD = Long_St / 60
End Function

Private Sub List3_Click()
    Find_Line List3.List(List3.ListIndex)
End Sub

Private Sub Map_CT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 72 And Shift = 7 Then Frame1.Visible = False
    If KeyCode = 83 And Shift = 7 Then Frame1.Visible = True
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
        Line2.X1 = 15 - Map_CT.Left
        Line2.X2 = Line2.X1 + 60
        Line2.Y1 = Map_FM.Height - Map_CT.Top - 15
        Line2.Y2 = Line2.Y1
        Label7.Left = 26 - Map_CT.Left
        Label7.Top = Map_FM.Height - Map_CT.Top - 35
    End If
    Label1.Caption = "E:-" & Change_JW(Int(X), 1) & "° N:" & Change_JW(Int(Y), 2) & "°"
End Sub

Private Function Change_JW(XXX As Long, XY_ID As Integer) As Single
    Select Case XY_ID
        Case 1: Change_JW = Int(((114.65 - 113.75) / Map_CT.Width) * XXX * 100 + 11375) / 100
        Case 2: Change_JW = Int(((22.92 - 22.4) / Map_CT.Height) * (Map_CT.Height - XXX) * 100 + 2240) / 100
    End Select
End Function

Private Sub Show_Loca(X As Long, Y As Long)
Dim TOM_X As Single
Dim TOM_Y As Single
    TOM_X = Map_FM.Width / 2 - X
    TOM_Y = Map_FM.Height / 2 - Y
    If TOM_X > 0 Then TOM_X = 0
    If TOM_Y > 0 Then TOM_Y = 0
    If TOM_X + Map_CT.Width < Map_FM.Width Then TOM_X = Map_FM.Width - Map_CT.Width
    If TOM_Y + Map_CT.Height < Map_FM.Height Then TOM_Y = Map_FM.Height - Map_CT.Height
    Map_CT.Left = TOM_X
    Map_CT.Top = TOM_Y
    Map_CT_X.Value = TOM_X
    Map_CT_Y.Value = TOM_Y
End Sub


Private Sub Show_Line()
Dim Str_Arr() As String
Dim Str_Arr_A() As String
Dim Str_Arr_B() As String
Dim Str_Arr_C() As String
Dim T_X As Long
Dim T_Y As Long
Dim i As Integer
Dim ii As Integer

Line_Cls
Select Case Draw_Str(0)
    Case "LINE_UP"
        Str_Arr = Split(Draw_Str(1), ">")
        Picture1.Visible = True
    Case "LINE_DOWN"
        Str_Arr = Split(Draw_Str(2), ">")
        Picture1.Visible = True
    Case "STATION"
        Str_Arr = Split(Draw_Str(1) & ">" & Draw_Str(2), ">")
End Select

For i = 0 To UBound(Str_Arr) - 1
    Str_Arr_A = Split(Str_Arr(i), ":")
    For ii = 0 To UBound(Str_Arr_A)
        Str_Arr_B = Split(Str_Arr_A(ii), ",")
        If Load_CI > -1 Then
            If Load_CI <> 0 Then Load Line_Cell(Load_CI)
            Line_Cell(Load_CI).X1 = T_X
            Line_Cell(Load_CI).Y1 = T_Y
            Line_Cell(Load_CI).X2 = Str_Arr_B(0)
            Line_Cell(Load_CI).Y2 = Str_Arr_B(1)
            Line_Cell(Load_CI).Visible = True
        End If
        T_X = Str_Arr_B(0)
        T_Y = Str_Arr_B(1)
        Load_CI = Load_CI + 1
    Next
    
    If Load_LCAB > -1 Then
        Str_Arr_C = Split(Str_Arr_A(0), ",")
        If Load_LCAB = 0 Then Show_Loca Int(Str_Arr_C(0)), Int(Str_Arr_C(1))
        If Load_LCAB <> 0 Then Load Shape_CS(Load_LCAB)
        Shape_CS(Load_LCAB).Top = Str_Arr_C(1) - 7
        Shape_CS(Load_LCAB).Left = Str_Arr_C(0) - 7
        Shape_CS(Load_LCAB).Visible = True
    End If
    
    Load_LCAB = Load_LCAB + 1
Next

Load_CI = Load_CI - 1
Load_LCAB = Load_LCAB - 1
For i = 0 To Load_LCAB
    Shape_CS(i).ZOrder (0)
Next

End Sub

Private Sub Map_CT_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
    If Shift = 1 Or Shift = 2 Then
        
        If Temp_X <> 0 And Temp_Y <> 0 Then
            Map_CT.Line (Temp_X, Temp_Y)-(X, Y)
        End If
        
        Temp_X = X
        Temp_Y = Y
        
        If List2.List(List1.ListIndex) = " " Then
                List2.List(List1.ListIndex) = X & "," & Y
            Else
                List2.List(List1.ListIndex) = List2.List(List1.ListIndex) & ":" & X & "," & Y
        End If
        
        If Shift = 2 Then
            List1.ListIndex = List1.ListIndex + 1
            List2.ListIndex = List1.ListIndex
            List2.List(List1.ListIndex) = X & "," & Y
        End If
        
        Text1.Text = ""
        
        For i = 0 To List1.ListIndex
            Text1.Text = Text1.Text & ">" & List2.List(i)
        Next
        
        Text1.Text = Right(Text1.Text, Len(Text1.Text) - 1)
    
        Text1.SelStart = Len(Text1.Text)
    
        If List1.ListIndex = List1.ListCount - 1 Then
            RS_Init
            Select Case Path_Mode
                Case 1
                    Conn.Execute "UPDATE Line_Info SET Path_ICED=" & True & ", Path_Up_Going='" & Text1.Text & "' WHERE Line_Num Like('" & List_Input(2).List(List_Input(2).ListIndex) & "')"
                Case 2
                    Conn.Execute "UPDATE Line_Info SET Path_ICED=" & True & ", Path_Down_Going='" & Text1.Text & "' WHERE Line_Num Like('" & List_Input(2).List(List_Input(2).ListIndex) & "')"
            End Select
            MsgBox "当前线路录入完毕！谢谢！"
        End If
    
    End If
End Sub

Private Sub Map_CT_x_Scroll()
    Map_CT.Left = Map_CT_X.Value
End Sub

Private Sub Map_CT_Y_Scroll()
    Map_CT.Top = Map_CT_Y.Value
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0
            Draw_Str(0) = "LINE_UP"
        Case 1
            Draw_Str(0) = "LINE_DOWN"
    End Select
    Show_Line
End Sub

Private Sub Text_Input_Change(Index As Integer)
Dim Que_Str As String
Dim SQL_Str As String
    Select Case Index
        Case 0, 1, 3
            If Index = 0 Or 1 Then
                If Text_Input(0).Text = Text_Input(1).Text Then
                    Select Case Index
                        Case 0
                            If Text_Input(0).Text <> "请输入出发地" Then
                                MsgBox "您输入的“出发地”和“目的地”一致，请重新输入“出发地”。", vbInformation
                                Input_Init 10
                                Exit Sub
                            End If
                        Case 1
                            If Text_Input(1).Text <> "请输入目的地" Then
                                MsgBox "您输入的“出发地”和“目的地”一致，请重新输入“目的地”。", vbInformation
                                Input_Init 11
                                Exit Sub
                            End If
                    End Select
                End If
                
            End If
            If Text_Input(Index).Text = "请输入出发地" Or Text_Input(Index).Text = "请输入目的地" Or Text_Input(Index).Text = "请输入站点名" Then
                Text_Input(Index).ForeColor = RGB(147, 147, 147)
            Else
                Text_Input(Index).ForeColor = vbBlack
                If TX_Ch = True Then
                    Que_Str = Text_Input(Index).Text
                    
                    
                    SQL_Str = "SELECT Station_Name FROM Stations_Info"
                    If Len(Que_Str) > 0 Then
                        SQL_Str = SQL_Str & " WHERE Station_Name LIKE '%" & Que_Str & "%'"
                    End If
                    SQL_Str = SQL_Str & " ORDER BY Station_Name"
                    
                    RS_Init
                    Rs.Open SQL_Str, Conn
                    
                    List_Input(Index).Clear
                    
                    Do Until Rs.EOF = True
                        List_Input(Index).AddItem Rs.Fields.Item(0)
                        Rs.MoveNext
                    Loop
                End If
            End If
        Case 2
            If Text_Input(Index).Text = "请输入路线编号" Then
                Text_Input(Index).ForeColor = RGB(147, 147, 147)
            Else
                Que_Str = Text_Input(Index).Text
                Text_Input(Index).ForeColor = vbBlack
    
                SQL_Str = "SELECT Line_Num FROM Line_Info"
                If Len(Que_Str) > 0 Then
                    SQL_Str = SQL_Str & " WHERE Line_Num LIKE '%" & Que_Str & "%'"
                End If
                SQL_Str = SQL_Str & " ORDER BY Line_Num"
                
                RS_Init
                Rs.Open SQL_Str, Conn
                
                List_Input(Index).Clear
                
                Do Until Rs.EOF = True
                    List_Input(Index).AddItem Rs.Fields.Item(0)
                    Rs.MoveNext
                Loop
                If List_Input(2).ListCount = 1 And Text_Input(2).Text = List_Input(2).List(0) Then
                    List_Input(2).ListIndex = 0
                    List_Input_Click 2
                End If
            End If
    
        Case 4
            If Text_Input(Index).Text = "请输入地名" Then
                Text_Input(Index).ForeColor = RGB(147, 147, 147)
            Else
                Que_Str = Text_Input(Index).Text
                Text_Input(Index).ForeColor = vbBlack
    
                SQL_Str = "SELECT Location FROM Locations_Info"
                If Len(Que_Str) > 0 Then
                    SQL_Str = SQL_Str & " WHERE Location LIKE '%" & Que_Str & "%'"
                End If
                SQL_Str = SQL_Str & " ORDER BY Location"
                
                RS_Init
                Rs.Open SQL_Str, Conn
                
                List_Input(Index).Clear
                
                Do Until Rs.EOF = True
                    List_Input(Index).AddItem Rs.Fields.Item(0)
                    Rs.MoveNext
                Loop
            End If
    End Select
End Sub

Private Sub Text_Input_GotFocus(Index As Integer)
    TX_Ch = True
    Select Case Index
        Case 0, 1
            If Text_Input(Index).Text = "请输入出发地" Or Text_Input(Index).Text = "请输入目的地" Then
                Text_Input(Index).Text = ""
            End If
        Case 2
            If Text_Input(Index).Text = "请输入路线编号" Then
                Text_Input(Index).Text = ""
            End If
        Case 3
            If Text_Input(Index).Text = "请输入站点名" Then
                Text_Input(Index).Text = ""
                List3.Clear
            End If
        Case 4
            If Text_Input(Index).Text = "请输入地名" Then
                Text_Input(Index).Text = ""
            End If
    End Select
End Sub

Private Sub Text_Input_LostFocus(Index As Integer)
    If Len(Text_Input(Index).Text) = 0 Then Input_Init Index
End Sub

Private Sub Input_Init(Index As Integer)
    Select Case Index
        Case 0
            Text_Input(0).Text = "请输入出发地"
        Case 1
            Text_Input(1).Text = "请输入目的地"
        Case 2
            Text_Input(2).Text = "请输入路线编号"
        Case 3
            Text_Input(3).Text = "请输入站点名"
            List3.Clear
        Case 4
            Text_Input(4).Text = "请输入地名"
        Case 9
            Text_Input(0).Text = "请输入出发地"
            Text_Input(1).Text = "请输入目的地"
            Text_Input(2).Text = "请输入路线编号"
            Text_Input(3).Text = "请输入站点名"
            Text_Input(4).Text = "请输入地名"
            List3.Clear
        Case 10
            Text_Input(0).Text = ""
        Case 11
            Text_Input(1).Text = ""
    End Select
End Sub


Private Sub Timer1_Timer()
    Count_Close = Count_Close + 1
    If Count_Close > 3 Then Picture5.Visible = False
End Sub
