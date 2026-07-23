use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Full Domain Model Simulation (E-Commerce Order Processor)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_domain_order_item_subtotal_calculation() {
    let src = r#"
Class OrderItem
    Public Property ProductId As String
    Public Property UnitPrice As Decimal
    Public Property Quantity As Integer

    Public ReadOnly Property Subtotal As Decimal
        Get
            Return UnitPrice * Quantity
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim item As New OrderItem With {.ProductId = "P1", .UnitPrice = 19.99D, .Quantity = 3}
        Console.WriteLine(item.Subtotal)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["59.97"]);
}

#[test]
fn test_vb_domain_customer_tier_discount_calculation() {
    let src = r#"
Enum CustomerTier
    Standard
    Silver
    Gold
End Enum

Class Customer
    Public Property Tier As CustomerTier = CustomerTier.Gold

    Public Function CalculateDiscount(total As Decimal) As Decimal
        Select Case Tier
            Case CustomerTier.Gold
                Return total * 0.2D
            Case CustomerTier.Silver
                Return total * 0.1D
            Case Else
                Return 0D
        End Select
    End Function
End Class

Module Program
    Sub Main()
        Dim cust As New Customer()
        Console.WriteLine(cust.CalculateDiscount(100D))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20.0"]);
}

#[test]
fn test_vb_domain_order_status_state_transitions() {
    let src = r#"
Imports System

Enum OrderStatus
    Created
    Paid
    Shipped
    Cancelled
End Enum

Class Order
    Public Property Status As OrderStatus = OrderStatus.Created

    Public Sub Pay()
        If Status <> OrderStatus.Created Then Throw New InvalidOperationException("Cannot pay order in status " & Status)
        Status = OrderStatus.Paid
    End Sub

    Public Sub Ship()
        If Status <> OrderStatus.Paid Then Throw New InvalidOperationException("Cannot ship unpaid order")
        Status = OrderStatus.Shipped
    End Sub
End Class

Module Program
    Sub Main()
        Dim ord As New Order()
        ord.Pay()
        ord.Ship()
        Console.WriteLine(ord.Status.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Shipped"]);
}

#[test]
fn test_vb_domain_invalid_state_transition_throws() {
    let src = r#"
Imports System

Enum OrderStatus
    Created
    Paid
    Shipped
End Enum

Class Order
    Public Property Status As OrderStatus = OrderStatus.Created
    Public Sub Ship()
        If Status <> OrderStatus.Paid Then Throw New InvalidOperationException("Cannot ship unpaid order")
        Status = OrderStatus.Shipped
    End Sub
End Class

Module Program
    Sub Main()
        Dim ord As New Order()
        Try
            ord.Ship() ' Cannot ship unpaid order!
        Catch ex As InvalidOperationException
            Console.WriteLine(ex.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Cannot ship unpaid order"]);
}

#[test]
fn test_vb_domain_shopping_cart_add_and_remove_items() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Linq

Class CartItem
    Public Property Id As String
    Public Property Price As Decimal
End Class

Class ShoppingCart
    Private items As New List(Of CartItem)()

    Public Sub Add(item As CartItem)
        items.Add(item)
    End Sub

    Public Sub Remove(id As String)
        items.RemoveAll(Function(i) i.Id = id)
    End Sub

    Public ReadOnly Property Total As Decimal
        Get
            Return items.Sum(Function(i) i.Price)
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim cart As New ShoppingCart()
        cart.Add(New CartItem With {.Id = "I1", .Price = 10D})
        cart.Add(New CartItem With {.Id = "I2", .Price = 20D})
        cart.Remove("I1")
        Console.WriteLine(cart.Total)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20"]);
}

#[test]
fn test_vb_domain_event_order_placed_notification() {
    let src = r#"
Imports System

Class OrderPlacedEventArgs
    Inherits EventArgs
    Public Property OrderId As String
End Class

Class OrderService
    Public Event OrderPlaced As EventHandler(Of OrderPlacedEventArgs)

    Public Sub PlaceOrder(id As String)
        RaiseEvent OrderPlaced(Me, New OrderPlacedEventArgs With {.OrderId = id})
    End Sub
End Class

Module Program
    Sub Main()
        Dim service As New OrderService()
        AddHandler service.OrderPlaced, Sub(s, e) Console.WriteLine("Notification Sent For Order: " & e.OrderId)
        service.PlaceOrder("ORD-999")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Notification Sent For Order: ORD-999"]);
}

#[test]
fn test_vb_domain_inventory_stock_deduction() {
    let src = r#"
Imports System

Class InventoryItem
    Public Property Sku As String
    Public Property QuantityOnHand As Integer

    Public Sub DeductStock(qty As Integer)
        If qty > QuantityOnHand Then Throw New InvalidOperationException("Insufficient stock for " & Sku)
        QuantityOnHand -= qty
    End Sub
End Class

Module Program
    Sub Main()
        Dim inv As New InventoryItem With {.Sku = "SKU-A", .QuantityOnHand = 50}
        inv.DeductStock(10)
        Console.WriteLine(inv.QuantityOnHand)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["40"]);
}

#[test]
fn test_vb_domain_payment_processor_strategy_pattern() {
    let src = r#"
Interface IPaymentStrategy
    Function ProcessPayment(amount As Decimal) As Boolean
End Interface

Class CreditCardPayment
    Implements IPaymentStrategy
    Public Function ProcessPayment(amount As Decimal) As Boolean Implements IPaymentStrategy.ProcessPayment
        Console.WriteLine("Paid $" & amount & " via CreditCard")
        Return True
    End Function
End Class

Class PayPalPayment
    Implements IPaymentStrategy
    Public Function ProcessPayment(amount As Decimal) As Boolean Implements IPaymentStrategy.ProcessPayment
        Console.WriteLine("Paid $" & amount & " via PayPal")
        Return True
    End Function
End Class

Module Program
    Sub Main()
        Dim strategy As IPaymentStrategy = New CreditCardPayment()
        strategy.ProcessPayment(45.5D)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Paid $45.5 via CreditCard"]);
}

#[test]
fn test_vb_domain_shipping_calculator_decorator_pattern() {
    let src = r#"
Interface IShippingCost
    Function CalculateCost() As Decimal
End Interface

Class BaseShipping
    Implements IShippingCost
    Public Function CalculateCost() As Decimal Implements IShippingCost.CalculateCost
        Return 5.0D
    End Function
End Class

Class ExpressShippingDecorator
    Implements IShippingCost
    Private baseCost As IShippingCost
    Public Sub New(inner As IShippingCost)
        baseCost = inner
    End Sub
    Public Function CalculateCost() As Decimal Implements IShippingCost.CalculateCost
        Return baseCost.CalculateCost() + 15.0D
    End Function
End Class

Module Program
    Sub Main()
        Dim cost As IShippingCost = New ExpressShippingDecorator(New BaseShipping())
        Console.WriteLine(cost.CalculateCost())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20.0"]);
}

#[test]
fn test_vb_domain_value_object_address_equality() {
    let src = r#"
Imports System

Class Address
    Implements IEquatable(Of Address)

    Public Property Street As String
    Public Property City As String
    Public Property Zip As String

    Public Function Equals1(other As Address) As Boolean Implements IEquatable(Of Address).Equals
        If other Is Nothing Then Return False
        Return Street = other.Street AndAlso City = other.City AndAlso Zip = other.Zip
    End Function
End Class

Module Program
    Sub Main()
        Dim a1 As New Address With {.Street = "123 Main St", .City = "NY", .Zip = "10001"}
        Dim a2 As New Address With {.Street = "123 Main St", .City = "NY", .Zip = "10001"}
        Console.WriteLine(a1.Equals1(a2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_domain_audit_log_immutable_record() {
    let src = r#"
Imports System

Structure AuditRecord
    Public ReadOnly Timestamp As DateTime
    Public ReadOnly Action As String
    Public ReadOnly User As String

    Public Sub New(act As String, u As String)
        Timestamp = New DateTime(2025, 1, 1)
        Action = act
        User = u
    End Sub
End Structure

Module Program
    Sub Main()
        Dim rec As New AuditRecord("LOGIN", "Admin")
        Console.WriteLine(rec.User & "|" & rec.Action)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Admin|LOGIN"]);
}

#[test]
fn test_vb_domain_coupon_code_validation() {
    let src = r#"
Class Coupon
    Public Property Code As String
    Public Property ExpiryYear As Integer
    Public Property MinOrderAmount As Decimal

    Public Function IsValid(year As Integer, total As Decimal) As Boolean
        Return year <= ExpiryYear AndAlso total >= MinOrderAmount
    End Function
End Class

Module Program
    Sub Main()
        Dim c As New Coupon With {.Code = "SAVE20", .ExpiryYear = 2026, .MinOrderAmount = 50D}
        Console.WriteLine(c.IsValid(2025, 75D) & "|" & c.IsValid(2025, 30D))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_domain_tax_calculator_by_region() {
    let src = r#"
Imports System.Collections.Generic

Class TaxCalculator
    Private taxRates As New Dictionary(Of String, Decimal) From {
        {"NY", 0.08D},
        {"CA", 0.10D},
        {"TX", 0.06D}
    }

    Public Function CalculateTax(state As String, amount As Decimal) As Decimal
        If taxRates.ContainsKey(state) Then
            Return amount * taxRates(state)
        End If
        Return 0D
    End Function
End Class

Module Program
    Sub Main()
        Dim calc As New TaxCalculator()
        Console.WriteLine(calc.CalculateTax("CA", 200D))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20.00"]);
}

#[test]
fn test_vb_domain_order_repository_in_memory() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Linq

Class OrderEntity
    Public Property OrderId As String
    Public Property CustomerName As String
End Class

Class OrderRepository
    Private db As New List(Of OrderEntity)()

    Public Sub Save(ord As OrderEntity)
        db.Add(ord)
    End Sub

    Public Function FindByCustomer(name As String) As List(Of OrderEntity)
        Return db.Where(Function(o) o.CustomerName = name).ToList()
    End Function
End Class

Module Program
    Sub Main()
        Dim repo As New OrderRepository()
        repo.Save(New OrderEntity With {.OrderId = "O1", .CustomerName = "Alice"})
        repo.Save(New OrderEntity With {.OrderId = "O2", .CustomerName = "Alice"})
        repo.Save(New OrderEntity With {.OrderId = "O3", .CustomerName = "Bob"})

        Console.WriteLine(repo.FindByCustomer("Alice").Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_domain_order_builder_fluent_interface() {
    let src = r#"
Imports System.Collections.Generic

Class FluentOrderBuilder
    Private custName As String
    Private items As New List(Of String)()

    Public Function WithCustomer(name As String) As FluentOrderBuilder
        custName = name
        Return Me
    End Function

    Public Function AddItem(item As String) As FluentOrderBuilder
        items.Add(item)
        Return Me
    End Function

    Public Function Build() As String
        Return custName & ":" & String.Join(",", items)
    End Function
End Class

Module Program
    Sub Main()
        Dim summary = New FluentOrderBuilder().WithCustomer("Charlie").AddItem("Laptop").AddItem("Mouse").Build()
        Console.WriteLine(summary)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Charlie:Laptop,Mouse"]);
}

#[test]
fn test_vb_domain_order_tracking_history() {
    let src = r#"
Imports System.Collections.Generic

Class TrackingEvent
    Public Status As String
    Public Location As String
    Public Sub New(s As String, l As String)
        Status = s
        Location = l
    End Sub
End Class

Module Program
    Sub Main()
        Dim history As New Stack(Of TrackingEvent)()
        history.Push(New TrackingEvent("Dispatched", "Warehouse A"))
        history.Push(New TrackingEvent("In Transit", "Hub B"))
        history.Push(New TrackingEvent("Out For Delivery", "Local Hub"))

        Dim current = history.Peek()
        Console.WriteLine(current.Status & " at " & current.Location)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Out For Delivery at Local Hub"]);
}

#[test]
fn test_vb_domain_invoice_pdf_generator_mock() {
    let src = r#"
Imports System.Text

Class InvoiceGenerator
    Public Function RenderInvoice(invoiceId As String, amount As Decimal) As String
        Dim sb As New StringBuilder()
        sb.AppendLine("=== INVOICE ===")
        sb.AppendLine("ID: " & invoiceId)
        sb.AppendLine("Amount: $" & amount)
        Return sb.ToString().Trim()
    End Function
End Class

Module Program
    Sub Main()
        Dim gen As New InvoiceGenerator()
        Dim txt = gen.RenderInvoice("INV-500", 150D)
        Console.WriteLine(txt.Contains("INV-500"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_domain_currency_exchange_converter() {
    let src = r#"
Imports System.Collections.Generic

Class CurrencyConverter
    Private rates As New Dictionary(Of String, Decimal) From {
        {"USD_EUR", 0.92D},
        {"USD_GBP", 0.78D}
    }

    Public Function Convert(amount As Decimal, pair As String) As Decimal
        Return amount * rates(pair)
    End Function
End Class

Module Program
    Sub Main()
        Dim cc As New CurrencyConverter()
        Console.WriteLine(cc.Convert(100D, "USD_EUR"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["92.00"]);
}

#[test]
fn test_vb_domain_order_batch_processor_concurrent_queue() {
    let src = r#"
Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim orderQueue As New ConcurrentQueue(Of String)()
        orderQueue.Enqueue("Ord1")
        orderQueue.Enqueue("Ord2")

        Dim processedCount = 0
        Dim id As String = Nothing
        While orderQueue.TryDequeue(id)
            processedCount += 1
        End While
        Console.WriteLine("Processed Orders: " & processedCount)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Processed Orders: 2"]);
}

#[test]
fn test_vb_domain_end_to_end_order_fulfillment_simulation() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Linq

Class LineItem
    Public Property Sku As String
    Public Property Price As Decimal
End Class

Class OrderAggregate
    Public Property Id As String
    Public Property Items As New List(Of LineItem)()
    Public Property IsPaid As Boolean = False
    Public Property IsShipped As Boolean = False

    Public Function CalculateTotal() As Decimal
        Return Items.Sum(Function(i) i.Price)
    End Function

    Public Sub ProcessFulfillment()
        If Not IsPaid Then Throw New System.InvalidOperationException("Unpaid")
        IsShipped = True
    End Sub
End Class

Module Program
    Sub Main()
        Dim ord As New OrderAggregate With {.Id = "E2E-100"}
        ord.Items.Add(New LineItem With {.Sku = "ITEM-1", .Price = 50D})
        ord.Items.Add(New LineItem With {.Sku = "ITEM-2", .Price = 50D})

        Dim total = ord.CalculateTotal()
        ord.IsPaid = (total = 100D)
        ord.ProcessFulfillment()

        Console.WriteLine("Order: " & ord.Id & "|Total: " & total & "|Shipped: " & ord.IsShipped)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Order: E2E-100|Total: 100|Shipped: True"]);
}
