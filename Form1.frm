VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9075
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   605
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "站点查询"
      Height          =   330
      Index           =   2
      Left            =   10665
      TabIndex        =   24
      Top             =   585
      Width           =   1320
   End
   Begin VB.CommandButton Command2 
      Caption         =   "线路查询"
      Height          =   330
      Index           =   1
      Left            =   9225
      TabIndex        =   23
      Top             =   585
      Width           =   1320
   End
   Begin VB.CommandButton Command2 
      Caption         =   "换乘指南"
      Height          =   330
      Index           =   0
      Left            =   7830
      TabIndex        =   22
      Top             =   585
      Width           =   1320
   End
   Begin VB.Frame Frame_CT 
      Caption         =   "Frame1"
      Height          =   4605
      Index           =   2
      Left            =   7830
      TabIndex        =   19
      Top             =   945
      Visible         =   0   'False
      Width           =   5280
      Begin VB.CommandButton Command1 
         Caption         =   "Check"
         Height          =   465
         Index           =   5
         Left            =   270
         TabIndex        =   28
         Top             =   3420
         Width           =   2265
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Reset"
         Height          =   465
         Index           =   4
         Left            =   2835
         TabIndex        =   27
         Top             =   3420
         Width           =   1860
      End
      Begin VB.ListBox List_Input 
         Height          =   2400
         Index           =   3
         Left            =   225
         TabIndex        =   21
         Top             =   855
         Width           =   1950
      End
      Begin VB.TextBox Text_Input 
         Height          =   375
         Index           =   3
         Left            =   225
         TabIndex        =   20
         Top             =   450
         Width           =   1995
      End
   End
   Begin VB.Frame Frame_CT 
      Caption         =   "Frame1"
      Height          =   4605
      Index           =   1
      Left            =   7830
      TabIndex        =   16
      Top             =   945
      Visible         =   0   'False
      Width           =   5280
      Begin VB.CommandButton Command1 
         Caption         =   "Check"
         Height          =   465
         Index           =   3
         Left            =   450
         TabIndex        =   26
         Top             =   3690
         Width           =   2265
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Reset"
         Height          =   465
         Index           =   2
         Left            =   3060
         TabIndex        =   25
         Top             =   3690
         Width           =   1860
      End
      Begin VB.TextBox Text_Input 
         Height          =   375
         Index           =   2
         Left            =   225
         TabIndex        =   18
         Top             =   450
         Width           =   1995
      End
      Begin VB.ListBox List_Input 
         Height          =   2400
         Index           =   2
         Left            =   225
         TabIndex        =   17
         Top             =   855
         Width           =   1950
      End
   End
   Begin VB.ListBox List_Loca 
      Height          =   6720
      Left            =   13140
      TabIndex        =   14
      Top             =   315
      Width           =   2490
   End
   Begin VB.ListBox List_Output 
      Height          =   1860
      ItemData        =   "Form1.frx":0000
      Left            =   270
      List            =   "Form1.frx":0002
      TabIndex        =   12
      Top             =   7155
      Width           =   14370
   End
   Begin VB.Frame Frame_CT 
      Caption         =   "Frame1"
      Height          =   4560
      Index           =   0
      Left            =   7830
      TabIndex        =   6
      Top             =   945
      Width           =   5190
      Begin VB.CommandButton Command1 
         Caption         =   "Reset"
         Height          =   465
         Index           =   1
         Left            =   3105
         TabIndex        =   13
         Top             =   3375
         Width           =   1860
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Check"
         Height          =   465
         Index           =   0
         Left            =   540
         TabIndex        =   11
         Top             =   3375
         Width           =   2265
      End
      Begin VB.ListBox List_Input 
         Height          =   2400
         Index           =   1
         Left            =   3060
         TabIndex        =   10
         Top             =   855
         Width           =   1950
      End
      Begin VB.TextBox Text_Input 
         Height          =   375
         Index           =   1
         Left            =   3060
         TabIndex        =   9
         Top             =   450
         Width           =   1995
      End
      Begin VB.TextBox Text_Input 
         Height          =   375
         Index           =   0
         Left            =   540
         TabIndex        =   8
         Top             =   450
         Width           =   2220
      End
      Begin VB.ListBox List_Input 
         Height          =   2400
         Index           =   0
         Left            =   540
         TabIndex        =   7
         Top             =   855
         Width           =   2220
      End
   End
   Begin VB.HScrollBar Map_CT_X 
      Height          =   280
      LargeChange     =   1000
      Left            =   90
      SmallChange     =   100
      TabIndex        =   5
      Top             =   5265
      Width           =   7305
   End
   Begin VB.VScrollBar Map_CT_Y 
      Height          =   5145
      LargeChange     =   1000
      Left            =   7425
      SmallChange     =   100
      TabIndex        =   4
      Top             =   90
      Width           =   280
   End
   Begin VB.PictureBox Map_FM 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5145
      Left            =   90
      ScaleHeight     =   341
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   485
      TabIndex        =   2
      Top             =   90
      Width           =   7305
      Begin VB.PictureBox Map_CT 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   54000
         Left            =   -405
         ScaleHeight     =   3600
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   5400
         TabIndex        =   3
         Top             =   540
         Width           =   81000
         Begin VB.Shape Shape_LC 
            BorderColor     =   &H000000C0&
            Height          =   1500
            Left            =   1395
            Shape           =   3  'Circle
            Top             =   810
            Visible         =   0   'False
            Width           =   1500
         End
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   1095
      Left            =   360
      TabIndex        =   15
      Top             =   5715
      Width           =   10725
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   240
      Index           =   1
      Left            =   9315
      TabIndex        =   1
      Top             =   90
      Width           =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   240
      Index           =   0
      Left            =   7830
      TabIndex        =   0
      Top             =   90
      Width           =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================
'地图显示模块 DEMO 7
'
'设计/编程: 黄思夏
'Design/Program: Leask Huang
'Copyfree (C) 2008 Leask Huang
'
'==============================================================

Option Explicit
    
    Dim Move_X As Long
    Dim Move_Y As Long
    
    Private Rs As New ADODB.Recordset
    Private Conn As New ADODB.Connection


Private Sub Map_Init()
    Map_CT.Picture = LoadPicture(App.Path & "\Library\Shenzhen\city_map.gif")
    Map_CT_X.Top = Map_FM.Top + Map_FM.Height + 2
    Map_CT_Y.Left = Map_FM.Left + Map_FM.Width + 2
    Map_CT_X.Height = 20
    Map_CT_Y.Width = 20
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



Private Sub Command1_Click(Index As Integer)
On Error Resume Next
Dim Do_i As Integer
Dim Do_ii As Integer
Dim i As Integer
Dim ii As Integer
Dim iii As Integer
Dim Result_Arr(99, 7) As String
Dim Result_Arr_i As Integer
Dim Line_Colt(99, 1) As String
Dim Str_SPL() As String
Dim Str_SPL_A() As String
Dim Str_i As Integer
Dim ST_OK As Boolean
Dim Find_Si As Integer
Dim Find_OK As Boolean

    Select Case Index
        Case 0
            Result_Arr_i = 0
            For Do_i = 0 To 1
                RS_Init
                Select Case Do_i
                    Case 0: Rs.Open "SELECT * FROM Line_Info WHERE Stations_Up_Going LIKE '%" & Text_Input(0).Text & "%" & Text_Input(1).Text & "%'", Conn
                    Case 1: Rs.Open "SELECT * FROM Line_Info WHERE Stations_Down_Going LIKE '%" & Text_Input(0).Text & "%" & Text_Input(1).Text & "%'", Conn
                End Select
                Do Until Rs.EOF = True
                    Result_Arr(Result_Arr_i, 0) = Result_Arr_i ' 方案编号
                    Result_Arr(Result_Arr_i, 1) = Rs.Fields.Item(1) '第一线路号码
                    Result_Arr(Result_Arr_i, 2) = "0" '第二线路号码
                    Select Case Do_i
                        Case 0: Str_SPL = Split(Rs.Fields.Item(7), ">")
                        Case 1: Str_SPL = Split(Rs.Fields.Item(8), ">")
                    End Select
                    ST_OK = False
                    For i = 0 To UBound(Str_SPL)
                        If InStr(Str_SPL(i), Text_Input(0).Text) > 0 Then ST_OK = True
                        If InStr(Str_SPL(i), Text_Input(1).Text) > 0 And ST_OK = True Then
                            Result_Arr(Result_Arr_i, 3) = Result_Arr(Result_Arr_i, 3) & Str_SPL(i) '第一线路路线
                            Exit For
                        End If
                        If ST_OK = True Then Result_Arr(Result_Arr_i, 3) = Result_Arr(Result_Arr_i, 3) & Str_SPL(i) & ">>"
                    Next
                    List_Output.AddItem Result_Arr(Result_Arr_i, 0) & "::" & Result_Arr(Result_Arr_i, 1) & "::" & Result_Arr(Result_Arr_i, 3)
                    Rs.MoveNext
                    Result_Arr_i = Result_Arr_i + 1
                Loop
            Next
           
            For Do_i = 0 To 1
                RS_Init
                Select Case Do_i
                    Case 0: Rs.Open "SELECT * FROM Line_Info WHERE (Stations_Up_Going LIKE '%" & Text_Input(0).Text & "%') and (Not Stations_Up_Going LIKE '%" & Text_Input(1).Text & "%')", Conn
                    Case 1: Rs.Open "SELECT * FROM Line_Info WHERE (Stations_Down_Going LIKE '%" & Text_Input(0).Text & "%') and (Not Stations_Down_Going LIKE '%" & Text_Input(1).Text & "%')", Conn
                End Select
                Str_i = 0
                Do Until Rs.EOF = True
                    Line_Colt(Str_i, 0) = Rs.Fields.Item(1)
                    Select Case Do_i
                        Case 0: Line_Colt(Str_i, 1) = Rs.Fields.Item(7)
                        Case 1: Line_Colt(Str_i, 1) = Rs.Fields.Item(8)
                    End Select
                    Str_i = Str_i + 1
                    Rs.MoveNext
                Loop
                For i = 0 To 99
                    If Line_Colt(i, 0) = "" Then Exit For
                    Str_SPL() = Split(Line_Colt(i, 1), ">")
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
                                        Find_OK = True
                                        Result_Arr(Result_Arr_i, 0) = Result_Arr_i
                                        Result_Arr(Result_Arr_i, 1) = Line_Colt(i, 0)
                                        Result_Arr(Result_Arr_i, 2) = Rs.Fields.Item(1)
                                        For iii = Find_Si To ii - 1
                                            Result_Arr(Result_Arr_i, 3) = Result_Arr(Result_Arr_i, 3) & Str_SPL(iii) & ">>"
                                        Next
                                        Result_Arr(Result_Arr_i, 3) = Result_Arr(Result_Arr_i, 3) & Str_SPL(ii)
                                        Select Case Do_ii
                                            Case 0: Str_SPL_A = Split(Rs.Fields.Item(7), ">")
                                            Case 1: Str_SPL_A = Split(Rs.Fields.Item(8), ">")
                                        End Select
                                        ST_OK = False
                                        For iii = 0 To UBound(Str_SPL_A)
                                            If InStr(Str_SPL_A(iii), Str_SPL(ii)) > 0 Then ST_OK = True
                                            If InStr(Str_SPL_A(iii), Text_Input(1).Text) > 0 And ST_OK = True Then
                                                Result_Arr(Result_Arr_i, 4) = Result_Arr(Result_Arr_i, 4) & Str_SPL_A(iii) '第二线路路线
                                                Exit For
                                            End If
                                            If ST_OK = True Then Result_Arr(Result_Arr_i, 4) = Result_Arr(Result_Arr_i, 4) & Str_SPL_A(iii) & ">>"
                                        Next
                                        List_Output.AddItem Result_Arr(Result_Arr_i, 0) & "::" & Result_Arr(Result_Arr_i, 1) & "::" & Result_Arr(Result_Arr_i, 2) & "::" & Result_Arr(Result_Arr_i, 3) & "::" & Result_Arr(Result_Arr_i, 4)
                                        Result_Arr_i = Result_Arr_i + 1
                                        If Result_Arr_i > 99 Then Exit Sub
                                        Rs.MoveNext
                                    Loop
                                Next
                            End If
                        End If
                    Next
                Next
            Next
            
            Case 1
                Input_Init 0
                Input_Init 1
                Load_Stations
            
            Case 2
                Input_Init 2
                Load_Line
            
            Case 3, 5
                RS_Init
                Select Case Index
                    Case 3: Rs.Open "SELECT * FROM Line_Info WHERE Line_Num LIKE '%" & Text_Input(2).Text & "%' Order By Line_Num", Conn
                    Case 5: Rs.Open "SELECT * FROM Line_Info WHERE (Stations_Up_Going LIKE '%" & Text_Input(3).Text & "%') or (Stations_Down_Going LIKE '%" & Text_Input(3).Text & "%')  Order By Line_Num", Conn
                End Select
                Do Until Rs.EOF = True
                    For i = 1 To 8
                        Result_Arr(Result_Arr_i, i - 1) = Rs.Fields.Item(i)
                    Next
                    List_Output.AddItem Result_Arr(Result_Arr_i, 0) & "::" & Result_Arr(Result_Arr_i, 1) & "::" & Result_Arr(Result_Arr_i, 3) & "::" & Result_Arr(Result_Arr_i, 4) & "::" & Result_Arr(Result_Arr_i, 5) & "::" & Result_Arr(Result_Arr_i, 6) & "::" & Result_Arr(Result_Arr_i, 7)
                    Rs.MoveNext
                    Result_Arr_i = Result_Arr_i + 1
                Loop
                
            Case 4
                Input_Init 3
                Load_Stations
                
    End Select
End Sub





Private Sub Command2_Click(Index As Integer)
Dim i As Integer
For i = 0 To 2
    Frame_CT(i).Visible = False
Next
Frame_CT(Index).Visible = True
End Sub

Private Sub Form_Load()
    Map_Init
    DB_Init
    Load_Stations
    Input_Init 9
    Load_Locations
    Load_Line
End Sub

Private Sub DB_Init()
    Dim strConn As String
    strConn = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\test.mdb;Persist Security Info=False"
    Conn.CursorLocation = adUseClient
    Conn.Open strConn
End Sub

Private Sub RS_Init()
    If Rs.State <> adStateClosed Then Rs.Close
    Rs.CursorType = adOpenKeyset
    Rs.LockType = adLockOptimistic
End Sub

Private Sub Load_Stations()
    RS_Init
    Rs.Open "SELECT Station_Name FROM Stations_Info ORDER BY Station_Name", Conn
    
    List_Input(0).Clear
    List_Input(1).Clear
    List_Input(3).Clear
    
    Do Until Rs.EOF = True
        List_Input(0).AddItem Rs.Fields.Item(0)
        List_Input(1).AddItem Rs.Fields.Item(0)
        List_Input(3).AddItem Rs.Fields.Item(0)
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
    Rs.Open "SELECT * FROM Locations_Info ORDER BY Location", Conn
    
    List_Loca.Clear
    
    Do Until Rs.EOF = True
        List_Loca.AddItem Rs.Fields.Item(1) & "/" & Rs.Fields.Item(2) & "," & Rs.Fields.Item(3)
        Map_CT.CurrentX = Rs.Fields.Item(2)
        Map_CT.CurrentY = Rs.Fields.Item(3)
        Map_CT.Line (Rs.Fields.Item(2) - 5, Rs.Fields.Item(3) - 5)-(Rs.Fields.Item(2) + 5, Rs.Fields.Item(3) + 5), , BF
        Rs.MoveNext
    Loop

End Sub

Private Sub List_Input_Click(Index As Integer)
    Text_Input(Index).Text = List_Input(Index).Text
End Sub

Private Sub List_Loca_Click()
Dim i As Integer
Dim Str As String
    Str = List_Loca.List(List_Loca.ListIndex)
    For i = 1 To Len(Str)
        If Mid(Str, i, 1) = "/" Then Str = Left(Str, i - 1)
    Next
      
    RS_Init
    Rs.Open "SELECT * FROM Locations_Info WHERE Location LIKE '" & Str & "'", Conn
    Show_Loca Rs.Fields.Item(2), Rs.Fields.Item(3)
    Shape_LC.Move Rs.Fields.Item(2) - 50, Rs.Fields.Item(3) - 50
    Shape_LC.Visible = True
End Sub

Private Sub List_Loca_KeyDown(KeyCode As Integer, Shift As Integer)

Dim i As Integer
Dim Str As String

    If KeyCode = 46 Then
      Str = List_Loca.List(List_Loca.ListIndex)
      For i = 1 To Len(Str)
          If Mid(Str, i, 1) = "/" Then Str = Left(Str, i - 1)
      Next

      Conn.Execute "DELETE * FROM Locations_Info WHERE Location LIKE '" & Str & "'"
      Load_Locations
    End If
End Sub

Private Sub List_Output_Click()
Label2.Caption = List_Output.List(List_Output.ListIndex)
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

Private Sub Show_Loca(X As Single, Y As Single)
Map_CT.Left = Map_FM.Width / 2 - X
Map_CT.Top = Map_FM.Height / 2 - Y
End Sub

Private Sub Map_CT_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Loca_Str As String
    If Shift = 1 Then
      Loca_Str = InputBox("Inout Name")
      If Len(Loca_Str) > 0 Then
      Conn.Execute "INSERT INTO Locations_Info(Location,Position_X,Position_Y) VALUES ('" & Loca_Str & "', '" & X & "','" & Y & "')"
      Load_Locations
      End If
    End If
End Sub

Private Sub Map_CT_x_Scroll()
    Map_CT.Left = Map_CT_X.Value
End Sub

Private Sub Map_CT_Y_Scroll()
    Map_CT.Top = Map_CT_Y.Value
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
                        MsgBox "您输入的""出发地""和""目的地""一致，请重新输入""出发地""。", vbInformation
                    Case 1
                        MsgBox "您输入的""出发地""和""目的地""一致，请重新输入""目的地""。", vbInformation
                End Select
                Input_Init Index
            End If
            Exit Sub
        End If
        If Text_Input(Index).Text = "请输入出发地" Or Text_Input(Index).Text = "请输入目的地" Or Text_Input(Index).Text = "请输入站点名" Then
            Text_Input(Index).ForeColor = RGB(147, 147, 147)
        Else
            Que_Str = Text_Input(Index).Text
            Text_Input(Index).ForeColor = vbBlack
            
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
        End If
End Select
End Sub

Private Sub Text_Input_GotFocus(Index As Integer)
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
    End Select
End Sub


