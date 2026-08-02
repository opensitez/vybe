' vybe-test: vb/vb_anonymous_type_array_projections/test_vb_anonymous_type_linq_join_projection
' origin: languages/vb/tests/vb/test_vb_anonymous_type_array_projections.rs

Imports System.Linq

Class Order
    Public Property OrderID As Integer
    Public Property CustomerID As Integer
    Public Sub New(o As Integer, c As Integer) : OrderID = o : CustomerID = c : End Sub
End Class

Class Customer
    Public Property CustomerID As Integer
    Public Property Name As String
    Public Sub New(c As Integer, n As String) : CustomerID = c : Name = n : End Sub
End Class

Module Program
    Sub Main()
        Dim orders = {New Order(1, 101), New Order(2, 102)}
        Dim customers = {New Customer(101, "Alice"), New Customer(102, "Bob")}

        Dim joined = From o In orders
                     Join c In customers On o.CustomerID Equals c.CustomerID
                     Select New With {.OrderID = o.OrderID, .CustomerName = c.Name}

        For Each item In joined
            Console.WriteLine("No." & item.OrderID & " " & item.CustomerName)
        Next
    End Sub
End Module
