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

Private SqlString As String

Private TableName As String

Private CreationTable As Boolean

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
    If SqlString = "" Then
       SqlString = Value
    Else
       SqlString = SqlString & ", " & Value
    End If
End Sub

Public Function Nullable() As DatabaseSchema
    SqlString = Left(SqlString, Len(SqlString) - 9)
    Set Nullable = Me
End Function

Public Function Default(valor) As DatabaseSchema
    SqlString = SqlString & " DEFAULT " & valor
    Set Default = Me
End Function

Public Function Unique() As DatabaseSchema
    SqlString = SqlString & " UNIQUE"
    Set Unique = Me
End Function

Public Function Create(Name As String) As DatabaseSchema
    TableName = Name
    CreationTable = True
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
    Dim SqlAlter As String
    
    SqlAlter = "ALTER TABLE " & TableName & _
                " ALTER COLUMN " & Field & " INTEGER" & _
                " CONSTRAINT FK_" & Field & _
                " REFERENCES " & OnTable & " (" & References & ")" & _
                " ON DELETE CASCADE"
                
    Debug.Print SqlAlter
    Connection.Execute SqlAlter, , adCmdText
End Sub

Public Sub Execute()
    Connection.Execute SqlStringCreateTable, , adCmdText
    Debug.Print SqlString
End Sub

Private Sub Class_Terminate()
    Debug.Print SqlString
    
    If CreationTable Then
        Connection.Execute SqlStringCreateTable, , adCmdText
    End If
End Sub

Private Function SqlStringCreateTable() As String
    AddIdAutoincrement
    AddCreatedUpdated
    SqlStringCreateTable = "CREATE Table " & TableName & " (" & SqlString & ")"
End Function

Private Sub AddIdAutoincrement()
    SqlString = "id AUTOINCREMENT PRIMARY KEY, " & SqlString
End Sub

Private Sub AddCreatedUpdated()
    SqlString = SqlString & ", created_at DATETIME DEFAULT NOW(), updated_at DATETIME DEFAULT NOW()"
End Sub

Private Function Connection() As ADODB.Connection
    Dim Con As New ADODB.Connection
    Dim DBFullName As String
    
    DBFullName = ThisWorkbook.Path & "\db.accdb"
    
    Con.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBFullName & ";"
    
    Set Connection = Con
End Function
