Attribute VB_Name = "TestSchema"

Private Sub Test_Squema()
    Dim Specs As New SpecSuite
    Dim Table As New Schema
    
    With Specs.It("Agrega un campo al sql")
        Table.FieldString "Name", 50
        Table.FieldInteger "Edad"
        
        .Expect(Table.sql).ToEqual "Name VARCHAR(50)"
    End With
    
    InlineRunner.RunSuite Specs
End Sub
