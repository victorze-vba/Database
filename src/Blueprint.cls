VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Blueprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sql As String
Private Cn As ADODB.Connection
    
Public Function FieldString(Name As String, Length As Integer) As Blueprint
    sql = sql & Name & " VARCHAR(" & Length & "), "
End Function

Public Function FieldDouble(Name) As Blueprint
    
End Function

Public Function FieldInteger(Name) As Blueprint
    sql = sql & Name & " VARCHAR(" & Length & "), "
End Function

Public Function Unique() As Blueprint

End Function

Public Function NotRequired() As Blueprint
    
End Function

Public Function Increments() As Blueprint

End Function

Public Sub Create(Table As String)
    sql = Left(sql, Len(sql) - 2)
    
    sql = "CREATE TABLE " & Table & " (" & sql & ")"
    
    Init
    Debug.Print sql
    Cn.Execute sql, , adCmdText
End Sub

Public Sub Init()
    Set Cn = New ADODB.Connection
    Dim sPath As String
    
    sPath = ThisWorkbook.Path & "\db.accdb"
    
    Set Cn = New ADODB.Connection
    Cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sPath & ";Jet OLEDB:Database Password=;"
End Sub

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()

End Sub
