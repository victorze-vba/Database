Attribute VB_Name = "DatabaseSchemaExample"
Option Explicit

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
End Sub

Sub AlterTables()
    Schema.Table("customer_orders").Foreing "customer_id", "id", "customers"
    Schema.Table("customer_orders").Foreing "product_id", "id", "products"
End Sub

Sub DropTables()
    Schema.Drop "customer_orders"
    Schema.Drop "customers"
    Schema.Drop "products"
End Sub
