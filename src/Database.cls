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
Const ConAccess As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\dev\github\vba-dev\vba-database\db.accdb;Jet OLEDB:Database Password=;"

Public Function OpenConection() As Boolean
    Set cn = New ADODB.Connection
    cn.Open ConAccess
End Function

Public Sub Insert(Data As Scripting.Dictionary)
    Dim k As Variant
    Dim Fields As String
    Dim Values As String

    Call OpenConection

    For Each k In Data.Keys
        Fields = Fields & ", " & k
        Values = Values & ", " & FormatValue(Data(k))
    Next k

    Fields = Right(Fields, Len(Fields) - 2)
    Values = Right(Values, Len(Values) - 2)

    query = "INSERT INTO Tabla (" & Fields & ") VALUES (" & Values & ")"
    
    Debug.Print query
    cn.Execute query

    Call ClouseConection
End Sub

Function FormatValue(Value As Variant)
    If IsDate(Value) Then
        FormatValue = "#" & Format(Value, "yyyy-mm-dd hh:nn:ss") & "#"
    ElseIf WorksheetFunction.IsText(Value) Then
        FormatValue = "'" & Value & "'"
    ElseIf IsNumeric(Value) Then
        FormatValue = Value
    End If
End Function

Public Sub ClouseConection()
    If Not rs Is Nothing Then
        rs.Close: Set rs = Nothing
    End If
    If Not cn Is Nothing Then
        cn.Close: Set cn = Nothing
    End If
End Sub
