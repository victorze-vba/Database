Attribute VB_Name = "TestDatabase"
Option Explicit

Sub TestInsert()
    Dim Data As New Scripting.Dictionary
    Dim db As New Database
    
    Data("nombre") = "Victor Hugo Zevallos"
    Data("edad") = 18.21
    Data("fecha") = "15/09/2017"
    
    db.Table("Tabla").Insert Data
End Sub

Sub TestGetAll()
    Dim Row As Scripting.Dictionary
    Dim Rows As New Collection
    Dim db As New Database
    Dim k As Variant
    
    Set Rows = db.Table("Tabla").GetAll()
    
    For Each Row In Rows
        For Each k In Row.Keys
            Debug.Print k, Row(k)
        Next k
    Next Row
End Sub

Sub TestGetWhere()
    Dim Row As Scripting.Dictionary
    Dim Rows As New Collection
    Dim db As New Database
    Dim k As Variant
    
    Set Rows = db.Table("Tabla").GetWhere("id = 1")
    
    For Each Row In Rows
        For Each k In Row.Keys
            Debug.Print k, Row(k)
        Next k
    Next Row
End Sub