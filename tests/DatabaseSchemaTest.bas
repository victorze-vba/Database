Attribute VB_Name = "DatabaseSchemaTest"
Option Explicit

Private Function Schema() As DatabaseSchema
    Dim InstanceDatabaseSchema As New DatabaseSchema
    Set Schema = InstanceDatabaseSchema
End Function

Private Sub Test()
    Dim Test As New VBAUnit
    Dim Result As String
    
    With Test.It("FieldString: one argument | default VARCHAR(255)")
        With Schema.Create("table_name", False)
            .FieldString "name"
            Result = .GetSqlString
        End With
        
        .AssertEquals "name VARCHAR(255) NOT NULL", Result
    End With
    
    With Test.It("FieldString: two arguments | VARCHAR(length)")
        With Schema.Create("table_name", False)
            .FieldString "name", 60
            Result = .GetSqlString
        End With
        
        .AssertEquals "name VARCHAR(60) NOT NULL", Result
    End With
    
    With Test.It("FieldInteger")
        With Schema.Create("table_name", False)
            .FieldInteger "age"
            Result = .GetSqlString
        End With
        
        .AssertEquals "age INTEGER NOT NULL", Result
    End With
    
    With Test.It("FieldDouble")
        With Schema.Create("table_name", False)
            .FieldDouble "price"
            Result = .GetSqlString
        End With
        
        .AssertEquals "price DOUBLE NOT NULL", Result
    End With
    
    With Test.It("FieldBoolean")
        With Schema.Create("table_name", False)
            .FieldBoolean "vip"
            Result = .GetSqlString
        End With
        
        .AssertEquals "vip BIT NOT NULL", Result
    End With
    
    With Test.It("FieldDate")
        With Schema.Create("table_name", False)
            .FieldDate "birthday"
            Result = .GetSqlString
        End With
        
        .AssertEquals "birthday DATE NOT NULL", Result
    End With
    
    With Test.It("Nullable")
        With Schema.Create("table_name", False)
            .FieldString("name").Nullable
            .FieldInteger("age").Nullable
            Result = .GetSqlString
        End With
        
        .AssertEquals "name VARCHAR(255), age INTEGER", Result
    End With
    
    With Test.It("Nullable and Default")
        With Schema.Create("table_name", False)
            .FieldString("name").Nullable.Default "foo"
            Result = .GetSqlString
        End With
        
        .AssertEquals "name VARCHAR(255) DEFAULT foo", Result
    End With
    
    With Test.It("Unique")
        With Schema.Create("table_name", False)
            .FieldString("name").Unique
            Result = .GetSqlString
        End With
        
        .AssertEquals "name VARCHAR(255) NOT NULL UNIQUE", Result
    End With
    
    With Test.It("Nullable and Unique")
        With Schema.Create("table_name", False)
            .FieldString("name").Nullable.Unique
            Result = .GetSqlString
        End With

        .AssertEquals "name VARCHAR(255) UNIQUE", Result
    End With
    
    With Test.It("multiple fields")
        With Schema.Create("table_name", False)
            .FieldString "name", 60
            .FieldInteger("age").Nullable
            .FieldDate("birthday").Nullable
            .FieldBoolean("VIP").Nullable.Default 1
            Result = .GetSqlString
        End With

        .AssertEquals "name VARCHAR(60) NOT NULL, age INTEGER, birthday DATE, VIP BIT DEFAULT 1", Result
    End With
End Sub
