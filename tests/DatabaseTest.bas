Attribute VB_Name = "DatabaseTest"
Option Explicit

Function DB() As Database
    Dim InstanceDatabase As New Database
    Set DB = InstanceDatabase
End Function

Private Sub Tests()
    Dim Test As New VBAUnit
    Dim Row As Scripting.Dictionary
    Dim RowCollection As Collection
    Dim Result As String

    With Test.It("Insert")
        Set Row = New Scripting.Dictionary
        Row.Add "name", "Matias" ' field,  value
        Row.Add "bill_date", DateSerial(2018, 9, 14)
        Row.Add "vip", False
        Row.Add "age", 30
        With DB.Table("clients", False)
            .Insert Row
            Result = .GetQuery
        End With
    
        .AssertEquals "INSERT INTO clients (name, bill_date, vip, age) VALUES ('Matias', #2018-09-14 00:00:00#, Falso, 30)", Result
    End With
    
    With Test.It("GetAll")
        With DB.Table("clients", False)
            .GetAll
            Result = .GetQuery
        End With
        
        .AssertEquals "SELECT * FROM clients", Result
    End With
    
    With Test.It("SelectFields.GetAll")
        With DB.Table("clients", False)
            .SelectFields("name", "age").GetAll
            Result = .GetQuery
        End With
        
        .AssertEquals "SELECT name, age FROM clients", Result
    End With
    
     With Test.It("Where: one argument | WHERE conditions...")
        With DB.Table("clients", False)
            .Where("year BETWEEN 2005 and 2010").GetAll
            Result = .GetQuery
        End With
        
        .AssertEquals "SELECT * FROM clients WHERE year BETWEEN 2005 and 2010", Result
    End With
    
    With Test.It("Where: two arguments | WHERE arg1 = arg2 ")
        With DB.Table("clients", False)
            .Where("name", "foo").GetAll
            Result = .GetQuery
        End With
        
        .AssertEquals "SELECT * FROM clients WHERE name = 'foo'", Result
    End With
    
    With Test.It("Where: three arguments | WHERE arg1 arg2 arg3")
        With DB.Table("clients", False)
            .Where("age", ">", 18).GetAll
            Result = .GetQuery
        End With
        
        .AssertEquals "SELECT * FROM clients WHERE age > 18", Result
    End With
    
    With Test.It("Where: three arguments | WHERE arg1 arg2 arg3")
        With DB.Table("clients", False)
            .Where("name", "LIKE", "T*").GetAll
            Result = .GetQuery
        End With
        
        .AssertEquals "SELECT * FROM clients WHERE name LIKE 'T*'", Result
    End With
    
    With Test.It("OrderBy: one argument | ORDER BY arg")
        With DB.Table("clients", False)
            .OrderBy("year").GetAll
            Result = .GetQuery
        End With
        
        .AssertEquals "SELECT * FROM clients ORDER BY year", Result
    End With
    
    With Test.It("OrderBy: two arguments | ORDER BY arg1, arg2")
        With DB.Table("clients", False)
            .OrderBy("year DESC", "month").GetAll
            Result = .GetQuery
        End With
        
        .AssertEquals "SELECT * FROM clients ORDER BY year DESC, month", Result
    End With
    
    With Test.It("Where and OrderBy")
        With DB.Table("clients", False)
            .Where "vip", 1
            .OrderBy("year DESC", "month").GetAll
            Result = .GetQuery
        End With
        
        .AssertEquals "SELECT * FROM clients WHERE vip = 1 ORDER BY year DESC, month", Result
    End With
    
    With Test.It("GroupBy: one argument | GROUP BY arg")
        With DB.Table("clients", False)
            .SelectFields "year", "COUNT(*) AS record_count"
            .GroupBy("year").GetAll

            Result = .GetQuery
        End With

        .AssertEquals "SELECT year, COUNT(*) AS record_count FROM clients GROUP BY year", Result
    End With
    
    With Test.It("GroupBy: two arguments | GROUP BY arg, arg")
        With DB.Table("clients", False)
            .SelectFields "year", "month", "COUNT(*) AS record_count"
            .GroupBy("year", "month").GetAll

            Result = .GetQuery
        End With

        .AssertEquals "SELECT year, month, COUNT(*) AS record_count FROM clients GROUP BY year, month", Result
    End With
    
    With Test.It("Having")
        With DB.Table("sales", False)
            .SelectFields "month", "SUM(price) as total_price"
            .GroupBy "month"
            .Having "SUM(price) > 3500"
            .GetAll
            
            Result = .GetQuery
        End With

        .AssertEquals "SELECT month, SUM(price) as total_price FROM sales GROUP BY month " & _
                      "HAVING SUM(price) > 3500", Result
    End With
    
    With Test.It("Delete")
        With DB.Table("sales", False)
            .Delete "price > 500"
            
            Result = .GetQuery
        End With

        .AssertEquals "DELETE FROM sales WHERE price > 500", Result
    End With

    With Test.It("DeleteId")
        With DB.Table("sales", False)
            .DeleteId 23
            
            Result = .GetQuery
        End With

        .AssertEquals "DELETE FROM sales WHERE id = 23", Result
    End With

    With Test.It("Update")
        With DB.Table("clients", False)
            Set Row = New Scripting.Dictionary
            Row("name") = "foo"
            Row("age") = 13
            .Update Row, "id = 1"
            
            Result = .GetQuery
        End With

        .AssertEquals "UPDATE clients SET name = 'foo', age = 13, updated_at = NOW() WHERE id = 1", Result
    End With
End Sub
