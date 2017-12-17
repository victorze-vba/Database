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

Private pTableName As String

Private cn As ADODB.Connection
    
Public Property Get sql() As String
    sql = pSql
End Property

Public Property Let sql(Value As String)
    If pSql = "" Then
        pSql = Value
    Else
        pSql = sql & ", " & Value
    End If
End Property

Public Property Let TableName(Name As String)
    pTableName = Name
End Property

Public Function FieldBoolean(Name As String) As Schema
    sql = Name & " BIT NOT NULL"
    
    Set FieldBoolean = Me
End Function

Public Function FieldString(Name As String, Optional Length As Integer = 255) As Schema
    sql = Name & " VARCHAR(" & Length & ") NOT NULL"
    
    Set FieldString = Me
End Function

Public Function FieldInteger(Name As String) As Schema
    sql = Name & " INTEGER NOT NULL"
    
    Set FieldInteger = Me
End Function

Public Function FieldDouble(Name As String) As Schema
    sql = Name & " DOUBLE NOT NULL"
    
    Set FieldDouble = Me
End Function

Public Function FieldDate(Name As String) As Schema
    sql = Name & " DATE NOT NULL"
    
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

Public Sub Foreing(Field As String, References As String, OnTable As String)
    Dim SqlAlter As String
    
    SqlAlter = "ALTER TABLE " & pTableName & _
                " ALTER COLUMN " & Field & " INTEGER" & _
                " CONSTRAINT FK_" & Field & _
                " REFERENCES " & OnTable & " (" & References & ")" & _
                " ON DELETE CASCADE"
    Debug.Print SqlAlter
    cn.Execute SqlAlter, , adCmdText
End Sub

Public Sub Create(Optional TimeStamp As Boolean = True)
    pSql = "id AUTOINCREMENT PRIMARY KEY, " & pSql

    If TimeStamp Then
        pSql = "CREATE TABLE " & pTableName & " (" & pSql & _
               ", created_at DATETIME DEFAULT NOW(), updated_at DATETIME DEFAULT NOW())"
    Else
        pSql = "CREATE TABLE " & pTableName & " (" & pSql & ")"
    End If

    cn.Execute pSql, , adCmdText
End Sub

Public Sub Drop(Table As String)
    sql = "DROP TABLE " & Table
    cn.Execute sql, , adCmdText
End Sub

Private Sub Class_Initialize()
    Dim sPath As String

    Set cn = New ADODB.Connection

    sPath = ThisWorkbook.Path & "\db.accdb"

    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sPath & ";Jet OLEDB:Database Password=;"
End Sub

Private Sub Class_Terminate()
    Debug.Print pSql
End Sub




