VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatabaseSchema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Query As String

Private TableName As String

Private ExecuteSql As Boolean

Public Function FieldString(Name As String, Optional Length As Integer = 255) As DatabaseSchema
    AddSqlString Name & " VARCHAR(" & Length & ") NOT NULL"
    Set FieldString = Me
End Function

Public Function FieldInteger(Name As String) As DatabaseSchema
    AddSqlString Name & " INTEGER NOT NULL"
    Set FieldInteger = Me
End Function

Public Function FieldDouble(Name As String) As DatabaseSchema
    AddSqlString Name & " DOUBLE NOT NULL"
    Set FieldDouble = Me
End Function

Public Function FieldBoolean(Name As String) As DatabaseSchema
    AddSqlString Name & " BIT NOT NULL"
    Set FieldBoolean = Me
End Function

Public Function FieldDate(Name As String) As DatabaseSchema
    AddSqlString Name & " DATE NOT NULL"
    Set FieldDate = Me
End Function

Private Sub AddSqlString(Value As String)
    If Query = "" Then
       Query = Value
    Else
       Query = Query & ", " & Value
    End If
End Sub

Public Function Nullable() As DatabaseSchema
    Query = Left(Query, Len(Query) - 9)
    Set Nullable = Me
End Function

Public Function Default(valor) As DatabaseSchema
    Query = Query & " DEFAULT " & valor
    Set Default = Me
End Function

Public Function Unique() As DatabaseSchema
    Query = Query & " UNIQUE"
    Set Unique = Me
End Function

Public Function Create(Name As String, Optional Execute As Boolean = True) As DatabaseSchema
    TableName = Name
    ExecuteSql = Execute
    Set Create = Me
End Function

Public Sub Drop(Name As String)
    Connection.Execute "DROP TABLE " & Name, , adCmdText
End Sub

Public Function Table(Name As String) As DatabaseSchema
    TableName = Name
    Set Table = Me
End Function

Public Sub Foreing(Field As String, References As String, OnTable As String)
    Query = "ALTER TABLE " & TableName & _
                " ADD CONSTRAINT FK_" & TableName & "_" & Field & "_" & OnTable & _
                " FOREIGN KEY (" & Field & ") REFERENCES " & _
                    OnTable & " (" & References & ")" & _
                " ON DELETE CASCADE"
    
    Debug.Print Query
    Connection.Execute Query, , adCmdText
End Sub

Private Sub Class_Terminate()
    If ExecuteSql Then
        Debug.Print SqlStringCreateTable
        Connection.Execute SqlStringCreateTable, , adCmdText
    End If
End Sub

Private Function SqlStringCreateTable() As String
    Dim Sql As String

    Sql = AddIdAutoincrement(Query)
    Sql = AddCreatedUpdated(Sql)

    SqlStringCreateTable = "CREATE TABLE " & TableName & " (" & Sql & ")"
End Function

Private Function AddIdAutoincrement(SqlStr As String) As String
    AddIdAutoincrement = "id AUTOINCREMENT PRIMARY KEY, " & SqlStr
End Function

Private Function AddCreatedUpdated(SqlStr As String) As String
    AddCreatedUpdated = SqlStr & ", created_at DATETIME DEFAULT NOW(), updated_at DATETIME DEFAULT NOW()"
End Function

Private Function Connection() As ADODB.Connection
    Dim Con As New ADODB.Connection
    Dim DBFullName As String
    
    DBFullName = ThisWorkbook.Path & "\db.accdb"
    
    Con.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBFullName & ";"
    
    Set Connection = Con
End Function

'Tests
Public Function GetSqlString() As String
    GetSqlString = Query
End Function
