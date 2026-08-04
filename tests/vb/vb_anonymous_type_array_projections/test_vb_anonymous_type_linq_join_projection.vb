' vybe-test: vb/vb_anonymous_type_array_projections/test_vb_anonymous_type_linq_join_projection
' origin: languages/vb/tests/vb/test_vb_anonymous_type_array_projections.rs

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
            __P(CStr("No." & item.OrderID & " " & item.CustomerName))
        Next
        __Check("No.1 Alice
No.2 Bob")
    End Sub
End Module
