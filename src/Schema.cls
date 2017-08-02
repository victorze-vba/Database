VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Schema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pSql As String
Private Cn As ADODB.Connection
    
Public Property Get sql() As String
    sql = pSql
End Property

Public Property Let sql(value As String)
    If pSql = "" Then
        pSql = value
    Else
        pSql = sql & ", " & value
    End If
End Property

Public Function FieldString(name As String, Length As Integer) As Schema
    sql = sql & name & " VARCHAR(" & Length & ")" 'TEXT
End Function

Public Function FieldInteger(name) As Schema
    sql = name & " INTEGER"
    Debug.Print sql
End Function

Public Function FieldDouble(name) As Schema
    
End Function

Public Function Unique() As Schema

End Function

Public Function NotRequired() As Schema
    
End Function

Public Function Increments() As Schema

End Function

Public Sub Create(Table As String)
    If sql = "" Then
        sql = "CREATE TABLE " & Table
    Else
        sql = Left(sql, Len(sql) - 2)
        sql = "CREATE TABLE " & Table & " (" & sql & ")"
    End If
    
    Init
    Debug.Print sql
    Cn.Execute sql, , adCmdText
End Sub

Public Sub Drop(Table As String)
    Init
    
    sql = "DROP TABLE " & Table
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

