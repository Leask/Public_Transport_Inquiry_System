VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16215
   LinkTopic       =   "Form1"
   ScaleHeight     =   605
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1081
   StartUpPosition =   3  '窗口缺省
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
   Begin VB.Frame Frame_A 
      Caption         =   "Frame1"
      Height          =   4560
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
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   240
      Index           =   1
      Left            =   7830
      TabIndex        =   1
      Top             =   540
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
'地图显示模块 DEMO 5
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
    Select Case Index
        Case 0
            RS_Init
            Rs.Open "SELECT * FROM Line_Info WHERE (Stations_Up_Going LIKE '%" & Text_Input(0).Text & "%" & Text_Input(1).Text & "%') or (Stations_Down_Going LIKE '%" & Text_Input(0).Text & "%" & Text_Input(1).Text & "%')", Conn
            Do Until Rs.EOF = True
                List_Output.AddItem Rs.Fields.Item(1) & "/" & Rs.Fields.Item(8)
                Rs.MoveNext
            Loop
        Case 1
            Input_Init 9
            Load_Stations
    End Select
End Sub

Private Sub Form_Load()
    Map_Init
    DB_Init
    Load_Stations
    Input_Init 9
    Load_Locations
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
    
    Do Until Rs.EOF = True
        List_Input(0).AddItem Rs.Fields.Item(0)
        List_Input(1).AddItem Rs.Fields.Item(0)
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
        'Map_CT.ForeColor = vbRed
        Map_CT.Line (Rs.Fields.Item(2) - 5, Rs.Fields.Item(3) - 5)-(Rs.Fields.Item(2) + 5, Rs.Fields.Item(3) + 5), , BF
        Rs.MoveNext
    Loop

End Sub

Private Sub List_Input_Click(Index As Integer)
    Text_Input(Index).Text = List_Input(Index).Text
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
    
    If Text_Input(Index).Text = "输入出发地" Or Text_Input(Index).Text = "输入目的地" Then
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
End Sub

Private Sub Text_Input_GotFocus(Index As Integer)
    If Text_Input(Index).Text = "输入出发地" Or Text_Input(Index).Text = "输入目的地" Then
        Text_Input(Index).Text = ""
    End If
End Sub

Private Sub Text_Input_LostFocus(Index As Integer)
    If Len(Text_Input(Index).Text) = 0 Then Input_Init Index
End Sub

Private Sub Input_Init(Index As Integer)
    Select Case Index
        Case 0
            Text_Input(0).Text = "输入出发地"
        Case 1
            Text_Input(1).Text = "输入目的地"
        Case 9
            Text_Input(0).Text = "输入出发地"
            Text_Input(1).Text = "输入目的地"
    End Select
End Sub


