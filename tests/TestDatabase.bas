Attribute VB_Name = "TestDatabase"
Option Explicit

Sub testInset()
    Dim Data As New Scripting.Dictionary
    Dim db As New Database
    
    Data("nombre") = "Victor Hugo"
    Data("edad") = 18.21
    Data("fecha") = "15/09/2017 01:01:00"
    
    db.Insert Data
End Sub


