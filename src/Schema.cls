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
    sql = name & " VARCHAR(" & Length & ") NOT NULL"
    
    Set FieldString = Me
End Function

Public Function FieldInteger(name As String) As Schema
    sql = name & " INTEGER NOT NULL"
    
    Set FieldInteger = Me
End Function

Public Function FieldDouble(name As String) As Schema
    sql = name & " DOUBLE NOT NULL"
    
    Set FieldDouble = Me
End Function

Public Function FieldDatetime(name As String) As Schema
    sql = name & " DATETIME NOT NULL"
    
    Set FieldDatetime = Me
End Function

Public Function FieldTime(name As String) As Schema
    sql = name & " TIME NOT NULL"
    
    Set FieldTime = Me
End Function

Public Function FieldDate(name As String) As Schema
    sql = name & " DATE NOT NULL"
    
    Set FieldDate = Me
End Function

Public Function Nullable() As Schema
    pSql = Left(pSql, Len(pSql) - 9)
    
    Set Nullable = Me
End Function

Public Function Default(valor) As Schema
    pSql = pSql & " DEFAULT " & valor
    
    Set Default = Me
End Function

Public Function Unique() As Schema
    pSql = pSql & " UNIQUE"
    
    Set Unique = Me
End Function

Public Sub Create(Table As String)
    If sql = "" Then
        sql = "CREATE TABLE " & Table
    Else
        pSql = "[id] AUTOINCREMENT PRIMARY KEY, " & pSql
        pSql = "CREATE TABLE " & Table & " (" & sql & ")"
    End If
    
    Init
    Debug.Print pSql
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
    Debug.Print sql
End Sub

