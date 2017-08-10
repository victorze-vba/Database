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

Private cn As ADODB.Connection

Private rs As ADODB.Recordset

Private query As String

Private pTable As String

Private Rows As New Collection

Private Row As Scripting.Dictionary

Public Function Table(TableName As String) As Database
    pTable = TableName
    
    Set Table = Me
End Function

Public Sub Insert(Data As Scripting.Dictionary)
    Dim k As Variant
    Dim Fields As String
    Dim Values As String

    For Each k In Data.Keys
        Fields = Fields & ", " & k
        Values = Values & ", " & FormatValue(Data(k))
    Next k

    Fields = Right(Fields, Len(Fields) - 2)
    Values = Right(Values, Len(Values) - 2)

    query = "INSERT INTO " & pTable & " (" & Fields & ") VALUES (" & Values & ")"

    cn.Execute query
End Sub

Public Function GetAll() As Collection
    query = "SELECT * FROM " & pTable
    
    Call GetData
    
    Set GetAll = Rows
End Function

'$user = DB::table('users')->where('name', 'John')->first();
Public Function GetWhere(Condition As String) As Collection
    query = "SELECT * FROM " & pTable & " WHERE " & Condition
    
    Call GetData
    
    Set GetWhere = Rows
End Function

Private Sub GetData()
    Dim c As Integer

    rs.Open query, cn, adOpenStatic, adLockOptimistic

    Do While Not rs.EOF
        Set Row = New Scripting.Dictionary

        For c = 0 To rs.Fields.Count - 1
            Row(rs.Fields(c).Name) = rs.Fields(c)
        Next c
        
        Rows.Add Row
        
        rs.MoveNext
    Loop
End Sub

'$users = DB::table('users')->count();

Private Function FormatValue(Value As Variant)
    If IsDate(Value) Then
        FormatValue = "#" & Format(Value, "yyyy-mm-dd hh:nn:ss") & "#"
    ElseIf WorksheetFunction.IsText(Value) Then
        FormatValue = "'" & Value & "'"
    ElseIf IsNumeric(Value) Then
        FormatValue = Value
    End If
End Function

Private Sub Class_Initialize()
    Dim sPath As String
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    sPath = ThisWorkbook.Path & "\db.accdb" ' path access
    
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sPath & ";Jet OLEDB:Database Password=;"
End Sub

Private Sub Class_Terminate()
    Debug.Print query
    
    If Not rs Is Nothing Then
        rs.Close: Set rs = Nothing
    End If
    If Not cn Is Nothing Then
        cn.Close: Set cn = Nothing
    End If
End Sub
