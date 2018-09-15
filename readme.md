## Create tables

```vb
Sub CreateTables()
    With Schema.Create("customers")
        .FieldString "name", 60
        .FieldString("region").Nullable
        .FieldString("street_address").Nullable
        .FieldString("city").Nullable
        .FieldString("state").Nullable
        .FieldInteger("zip").Nullable
    End With

    With Schema.Create("customer_orders")
        .FieldDate "order_date"
        .FieldDate("ship_date").Nullable
        .FieldInteger "customer_id"
        .FieldInteger "product_id"
        .FieldInteger "order_qty"
        .FieldBoolean "shipped"
    End With

    With Schema.Create("products")
        .FieldString "description"
        .FieldDouble "price"
    End With

    ' foreing key
    Schema.Table("customer_orders").Foreing "customer_id", "id", "customers"
    Schema.Table("customer_orders").Foreing "product_id", "id", "products"
End Sub
```

![imagen](https://raw.githubusercontent.com/vba-dev/database/master/relations.png)


## Query Builder

```vb
sub InsertData
    Dim Row As New Scripting.Dictionary

    Row("description") = "Copper"
    Row("price") = 7.51

    DB.Table("products", False).Insert Row
    ' query:
    ' INSERT INTO products (description, price) VALUES ('Copper', 7.51)
end sub

Sub WhereAndOrderBy()
    Dim Products As Collection

    Set Products = DB.Table("products").Where("price", ">", 10).OrderBy("price DESC").GetAll
    ' query:
    ' SELECT * FROM products WHERE price > 10 ORDER BY price DESC
End SuInnerJoinCustomerAndCustomerb

Sub InnerJoinCustomerAndCustomerOrders()
    Dim Orders As Collection

    With DB.Table("customers")
        .SelectFields "customers.name", "ship_date", "order_qty"
        .Join "customer_orders", "customers.id", "=", "customer_orders.customer_id"
        Set Orders = .GetAll
    End With
    ' query:
    ' SELECT customers.name, ship_date, order_qty
    ' FROM customers
    ' INNER JOIN customer_orders
    ' ON customers.id = customer_orders.customer_id
End Sub
```


[More examples](https://github.com/vba-dev/database/blob/master/tests/examples/DatabaseExample.bas)
