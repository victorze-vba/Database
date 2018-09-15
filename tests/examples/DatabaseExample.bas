Attribute VB_Name = "DatabaseExample"
Option Explicit

Sub InsertCustomersTable()
    Dim Row1 As New Scripting.Dictionary
    Dim Row2 As New Scripting.Dictionary
    Dim Row3 As New Scripting.Dictionary
    Dim Row4 As New Scripting.Dictionary
    Dim Row5 As New Scripting.Dictionary
    
    Row1("name") = "LITE Industrial"
    Row1("region") = "Southwest"
    Row1("street_address") = "729 Ravine Way"
    Row1("city") = "Irving"
    Row1("state") = "TX"
    Row1("zip") = 75014
    
    Row2("name") = "Rex Tooling Inc"
    Row2("region") = "Southwest"
    Row2("street_address") = "6129 Collie Blvd"
    Row2("city") = "Dallas"
    Row2("state") = "TX"
    Row2("zip") = 75201
    
    Row3("name") = "Re-Barre Construction"
    Row3("region") = "Southwest"
    Row3("street_address") = "9043 Windy Dr"
    Row3("city") = "Irving"
    Row3("state") = "TX"
    Row3("zip") = 75032
    
    Row4("name") = "Prairie Construction"
    Row4("region") = "Southwest"
    Row4("street_address") = "264 Long Rd"
    Row4("city") = "Moore"
    Row4("state") = "OK"
    Row4("zip") = 62104
    
    Row5("name") = "Marsh Lane Metal Works"
    Row5("region") = "Southwest"
    Row5("street_address") = "9143 Marsh Ln"
    Row5("city") = "Avondale"
    Row5("state") = "LA"
    Row5("zip") = 79782
    
    DB.Table("customers").Insert Row1
    DB.Table("customers").Insert Row2
    DB.Table("customers").Insert Row3
    DB.Table("customers").Insert Row4
    DB.Table("customers").Insert Row5
End Sub

Sub InsertProductsTable()
    Dim Row1 As New Scripting.Dictionary
    Dim Row2 As New Scripting.Dictionary
    Dim Row3 As New Scripting.Dictionary
    Dim Row4 As New Scripting.Dictionary
    Dim Row5 As New Scripting.Dictionary
    Dim Row6 As New Scripting.Dictionary
    Dim Row7 As New Scripting.Dictionary
    Dim Row8 As New Scripting.Dictionary
    Dim Row9 As New Scripting.Dictionary
    
    Row1("description") = "Copper"
    Row1("price") = 7.51
    
    Row2("description") = "Aluminum"
    Row2("price") = 2.58
    
    Row3("description") = "Silver"
    Row3("price") = 15
    
    Row4("description") = "Steel"
    Row4("price") = 12.31
    
    Row5("description") = "Bronze"
    Row5("price") = 4
    
    Row6("description") = "Duralumin"
    Row6("price") = 7.6
    
    Row7("description") = "Solder"
    Row7("price") = 14.16
    
    Row8("description") = "Stellite"
    Row8("price") = 13.31
    
    Row9("description") = "Brass"
    Row9("price") = 4.75
    
    DB.Table("products").Insert Row1
    DB.Table("products").Insert Row2
    DB.Table("products").Insert Row3
    DB.Table("products").Insert Row4
    DB.Table("products").Insert Row5
    DB.Table("products").Insert Row6
    DB.Table("products").Insert Row7
    DB.Table("products").Insert Row8
    DB.Table("products").Insert Row9
End Sub

Sub InsertCustomerOrdersTable()
    Dim Row1 As New Scripting.Dictionary
    Dim Row2 As New Scripting.Dictionary
    Dim Row3 As New Scripting.Dictionary
    Dim Row4 As New Scripting.Dictionary
    Dim Row5 As New Scripting.Dictionary
    
    Row1("order_date") = DateSerial(2015, 5, 15)
    Row1("ship_date") = DateSerial(2015, 5, 18)
    Row1("customer_id") = 1
    Row1("product_id") = 1
    Row1("order_qty") = 450
    Row1("shipped") = False
    
    Row2("order_date") = DateSerial(2015, 5, 18)
    Row2("ship_date") = DateSerial(2015, 5, 21)
    Row2("customer_id") = 3
    Row2("product_id") = 2
    Row2("order_qty") = 600
    Row2("shipped") = False
    
    Row3("order_date") = DateSerial(2015, 5, 20)
    Row3("ship_date") = DateSerial(2015, 5, 23)
    Row3("customer_id") = 3
    Row3("product_id") = 5
    Row3("order_qty") = 300
    Row3("shipped") = False
    
    Row4("order_date") = DateSerial(2015, 5, 18)
    Row4("ship_date") = DateSerial(2015, 5, 22)
    Row4("customer_id") = 5
    Row4("product_id") = 4
    Row4("order_qty") = 375
    Row4("shipped") = False
    
    Row5("order_date") = DateSerial(2015, 5, 17)
    Row5("ship_date") = DateSerial(2015, 5, 20)
    Row5("customer_id") = 3
    Row5("product_id") = 2
    Row5("order_qty") = 500
    Row5("shipped") = False
    
    DB.Table("customer_orders").Insert Row1
    DB.Table("customer_orders").Insert Row2
    DB.Table("customer_orders").Insert Row3
    DB.Table("customer_orders").Insert Row4
    DB.Table("customer_orders").Insert Row5
End Sub

Sub InsertDataxx()
    Dim Data As Scripting.Dictionary
    Dim i As Integer
    
    For i = 1 To 100
        Set Data = New Scripting.Dictionary
        
        Data("client") = "client" & i
        Data("price") = i * 100
        
        DB.Table("sales").Insert Data
    Next i
End Sub

Sub GetAllProducts()
    Dim Products As Collection

    Set Products = DB.Table("products").SelectFields("description", "price").GetAll
    
    PrintCollection Products
End Sub

Sub JoinCustomerAndCustomerOrders()
    Dim Orders As Collection
    
    With DB.Table("customers")
        .SelectFields "customers.name", "ship_date", "order_qty"
        .Join "customer_orders", "customers.id", "=", "customer_orders.customer_id"
        Set Orders = .GetAll
    End With
    
    PrintCollection Orders
End Sub

Sub JoinCustomerAndCustomerOrdersWhereOrderBy()
    Dim Orders As Collection

    With DB.Table("customers")
        .SelectFields "name", "ship_date", "order_qty"
        .Join "customer_orders", "customers.id", "=", "customer_orders.customer_id"
        .Where "name", "LIKE", "Re%"
        .OrderBy "ship_date"
        Set Orders = .GetAll
    End With

    PrintCollection Orders
End Sub

Sub JoinNested()
    Dim Orders As Collection

    With DB.Table("customers")
        .SelectFields "name", "ship_date", "order_qty", "description", "price"
        .Join "customer_orders", "customers.id", "=", "customer_orders.customer_id"
        .Join "products", "customer_orders.product_id", "=", "products.id"
        .OrderBy "price"
        Set Orders = .GetAll
    End With

    PrintCollection Orders
End Sub

Sub PrintCollection(DataCollection As Collection)
    Dim DataDic As Scripting.Dictionary
    Dim Row As String
    Dim key As Variant
    
    For Each DataDic In DataCollection
        Row = ""
        For Each key In DataDic.Keys
            Row = Row & DataDic(key) & "|"
        Next key
        Debug.Print Row
    Next DataDic
End Sub

