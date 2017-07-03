Attribute VB_Name = "migrations"
Option Explicit

Sub UpUsers()
    Dim Table As New Blueprint
    
    Table.FieldString "nota", 50
    'Table.FieldInteger "edad"
    
    Table.Create "notas"
    
    
End Sub

Sub DownUsers()
    
End Sub
