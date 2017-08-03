Attribute VB_Name = "TestSchema"

Private Sub Test_Squema()
    Dim Specs As New SpecSuite
    Dim Table As New Schema
    
    With Specs.It("Agrega un campo al sql")
        Table.FieldString("Name", 50).Nullable
        Table.FieldInteger "Edad"
        Table.FieldDouble "Salario"
        
        
        .Expect(Table.sql).ToEqual "Name VARCHAR(50), Edad INTEGER"
    End With
    
    InlineRunner.RunSuite Specs
End Sub

Private Sub CreateTableTest()
    Dim Table As New Schema
    
    Table.FieldString("Name", 50).Unique
    Table.FieldInteger "Edad"
    Table.FieldDouble("Salario").Nullable.Default 153.33
    Table.FieldDate "Fecha"
    Table.FieldTime "HoraSalida"
    Table.FieldDatetime "create_at"
    
    Table.Create "TestTable"
End Sub

Private Sub DropTableTest()
    Dim Table As New Schema
    Table.Drop ("TestTable")
End Sub
