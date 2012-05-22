VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    

    
 ' Private Sub Form_Load()

   

  '    MsgBox Rs.RecordCount
      
   '   MsgBox Rs.Fields.Item(1)
    '  Rs.MoveNext
     ' MsgBox Rs.Fields.Item(1)
    '   绑定进DataGrid
   ' Set DataGrid1.DataSource = Rs
   
 ' Conn.Execute "INSERT INTO People(USName,USDONE,LOVE) VALUES ('LLLSSSSKKKKKK', 'hi there','LoveLoveLove')"
 ' End Sub
'Private Sub Command1_Click()
'On Error Resume Next
'Dim Sp_Arr() As String
'Dim i As Integer
'Dim ii As Integer
'Dim T_Str As Variant
'
' If Rs.State <> adStateClosed Then Rs.Close
'    Rs.CursorType = adOpenKeyset
'    Rs.LockType = adLockOptimistic
'
'Rs.Open "Select Stations_Up_Going,Stations_Down_Going from Line_Info", Conn
'Do Until Rs.EOF = True
'    For i = 0 To 1
'        Sp_Arr = Split(Rs.Fields.Item(i), ">")
'            For Each T_Str In Sp_Arr
'            Conn.Execute "INSERT INTO Stations_Info(Station_Name) VALUES ('" & T_Str & "')"
'            Next
'            Rs.MoveNext
'    Next
'Loop
'End Sub
