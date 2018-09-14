VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private TableName As String

Private Query As String

Private ExecuteQuery As Boolean

Private FieldsStr As String

Private WhereStr As String

Private OrderByStr As String

Private GroupByStr As String

Private HavingStr As String

Public Function SelectFields(ParamArray Fields() As Variant) As Database
    Dim Field As Variant

    For Each Field In Fields
        FieldsStr = IIf(FieldsStr = "", Field, FieldsStr & ", " & Field)
    Next Field

    Set SelectFields = Me
End Function

Public Function GetAll() As Collection
    Query = "SELECT " & GetFields & " FROM " & TableName

    AddWhere
    AddGroupBy
    AddHaving
    AddOrderBy

    If ExecuteQuery Then Set GetAll = GetData
End Function

Public Function Where(ParamArray Chunk() As Variant) As Database
    Dim NumberArgs As Integer
    NumberArgs = UBound(Chunk) + 1

    Select Case NumberArgs
        Case 1:
            WhereStr = Chunk(0)
        Case 2:
            WhereStr = Chunk(0) & " = " & FormatValue(Chunk(1))
        Case 3:
            WhereStr = Chunk(0) & " " & Chunk(1) & " " & FormatValue(Chunk(2))
    End Select

    Set Where = Me
End Function

Private Sub AddWhere()
    If WhereStr <> "" Then
        Query = Query & " WHERE " & WhereStr
    End If
End Sub

Private Sub AddOrderBy()
    If OrderByStr <> "" Then
        Query = Query & " ORDER BY " & OrderByStr
    End If
End Sub

Private Sub AddGroupBy()
    If GroupByStr <> "" Then
        Query = Query & " GROUP BY " & GroupByStr
    End If
End Sub

Private Sub AddHaving()
    If HavingStr <> "" Then
        Query = Query & " HAVING " & HavingStr
    End If
End Sub

Public Function OrderBy(ParamArray Fields() As Variant) As Database
    Dim Field As Variant

    For Each Field In Fields
        OrderByStr = IIf(OrderByStr = "", Field, OrderByStr & ", " & Field)
    Next Field

    Set OrderBy = Me
End Function

Public Function GroupBy(ParamArray Fields() As Variant) As Database
    Dim Field As Variant

    For Each Field In Fields
        GroupByStr = IIf(GroupByStr = "", Field, GroupByStr & ", " & Field)
    Next Field

    Set GroupBy = Me
End Function

Public Function Having(Condition As String) As Database
    HavingStr = Condition

    Set Having = Me
End Function

Private Function GetData() As Collection
    Dim column As Long
    Dim rs As ADODB.Recordset
    Dim Rows As New Collection
    Dim Row As Scripting.Dictionary

    Set rs = Recordset

    Do While Not rs.EOF
        Set Row = New Scripting.Dictionary

        For column = 0 To rs.Fields.Count - 1
            Row(rs.Fields(column).Name) = rs.Fields(column)
        Next column

        Rows.Add Row

        rs.MoveNext
    Loop

    Set GetData = Rows
End Function

Private Function GetFields() As String
    If FieldsStr = "" Then
        GetFields = "*"
    Else
        GetFields = FieldsStr
    End If
End Function

Public Function Table(Name As String, Optional Execute As Boolean = True) As Database
    TableName = Name
    ExecuteQuery = Execute
    Set Table = Me
End Function

Public Sub Insert(Data As Scripting.Dictionary)
    Dim Key As Variant
    Dim Fields As String
    Dim Values As String
    Dim Value As Variant

    For Each Key In Data.Keys
        Fields = IIf(Fields = "", Key, Fields & ", " & Key) ' field1, field2, ...
        Value = FormatValue(Data(Key))
        Values = IIf(Values = "", Value, Values & ", " & Value) ' value1, value2, ...
    Next Key

    Query = "INSERT INTO " & TableName & " (" & Fields & ") VALUES (" & Values & ")"

    ConnectionExecute
End Sub

Sub Update(Data As Scripting.Dictionary, Condition As String)
    Dim Key As Variant
    Dim FieldsValues As String
    Dim FieldValue As String

    For Each Key In Data.Keys
        FieldValue = Key & " = " & FormatValue(Data(Key)) ' field_name = value
        FieldsValues = IIf(FieldsValues = "", FieldValue, FieldsValues & ", " & FieldValue)
    Next Key

    FieldsValues = FieldsValues & ", updated_at = NOW()"

    Query = "UPDATE " & TableName & " SET " & FieldsValues & " WHERE " & Condition

    If ExecuteQuery Then Connection.Execute Query
End Sub

Sub Delete(Condition As String)
    Query = "DELETE FROM " & TableName & " WHERE " & Condition
    If ExecuteQuery Then Connection.Execute Query
End Sub

Sub DeleteId(Id As Integer)
    Query = "DELETE FROM " & TableName & " WHERE id = " & Id
    If ExecuteQuery Then Connection.Execute Query
End Sub

Sub Truncate()
    Query = "DELETE * FROM " & TableName
    If ExecuteQuery Then Connection.Execute Query
End Sub

Public Function Count() As Long
    Query = "SELECT id FROM " & TableName

    rs.Open Query, cn, adOpenStatic, adLockOptimistic

    Count = rs.RecordCount
End Function

Private Function FormatValue(Value As Variant)
    If IsDate(Value) Then
        FormatValue = "#" & Format(Value, "yyyy-mm-dd hh:nn:ss") & "#"
    ElseIf WorksheetFunction.IsText(Value) Then
        FormatValue = "'" & Value & "'"
    ElseIf IsNumeric(Value) Then
        FormatValue = Value
    End If
End Function

Private Sub Class_Terminate()
    Debug.Print Query
End Sub

Private Function Recordset() As ADODB.Recordset
    Dim rs As New ADODB.Recordset

    rs.Open Query, Connection, adOpenStatic, adLockOptimistic

    Set Recordset = rs
End Function

Private Function Connection() As ADODB.Connection
    Dim cn As New ADODB.Connection
    Dim DBFullName As String

    DBFullName = ThisWorkbook.Path & "\db.accdb"

    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBFullName & ";"

    Set Connection = cn
End Function

Private Sub ConnectionExecute()
    If ExecuteQuery Then
        Connection.Execute Query, , adCmdText
    End If
End Sub

'Tests
Public Function GetQuery() As String
    GetQuery = Query
End Function

