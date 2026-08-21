' vybe-test: vb/vb_full_domain_model_simulation/test_vb_domain_end_to_end_order_fulfillment_simulation
' origin: languages/vb/tests/vb/test_vb_full_domain_model_simulation.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.
'
' Output is COLLECTED, not paired. The emitter rewrites every
' `Console.WriteLine(x)` into `__P(CStr(x))` and compares the whole output once
' at the end of `Sub Main`. Pairing the i-th print with the i-th expected line
' cannot assert anything about a loop, and loops alone were 402 of VB's 6,671
' cases.
'
' Rendering happens at the CALL SITE via `CStr`, where the expression still has
' its static type — the same reason the C# harness renders with `.ToString()`
' rather than inside the helper.

Imports System.Collections.Generic
Imports System.Linq
Module VybeCheck
    Public __buf As String = ""

    Sub __P(s As String)
        __buf = __buf & s & vbLf
    End Sub

    Sub __Pr(s As String)
        __buf = __buf & s
    End Sub

    ' The final WriteLine contributes a trailing newline that the expected line
    ' vector never carried, so BOTH forms are accepted.
    Sub __Check(want As String)
        If __buf <> want AndAlso __buf <> want & vbLf Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & __buf & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module


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

        __P(CStr("Order: " & ord.Id & "|Total: " & total & "|Shipped: " & ord.IsShipped))
        __Check("Order: E2E-100|Total: 100|Shipped: True")
    End Sub
End Module
