' vybe-test: vb/vb_system_linq_join_matrix/linq_left_join_style_with_default_if_empty
' origin: languages/vb/tests/vb/test_vb_system_linq_join_matrix.rs

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

Class User
    Public Id As Integer
    Public Name As String
    Public Sub New(id As Integer, name As String)
        Me.Id = id
        Me.Name = name
    End Sub
End Class

Class Order
    Public UserId As Integer
    Public Amount As Integer
    Public Sub New(id As Integer, amount As Integer)
        Me.UserId = id
        Me.Amount = amount
    End Sub
End Class

Module M
    Sub Main()
        Dim users = {New User(1, "Ada"), New User(2, "Bob")}
        Dim orders = {New Order(1, 10)}

        Dim grouped = From u In users _
            Group Join o In orders On u.Id Equals o.UserId _
            Into g = Group _
            Select User = u.Name, Total = g.Sum(Function(x) x.Amount)

        Dim rows As Integer = grouped.Count()
        Dim firstTotal As Integer = grouped(0).Total
        Dim secondTotal As Integer = grouped(1).Total

        __P(CStr(rows))
        __P(CStr(firstTotal))
        __P(CStr(secondTotal))
        __Check("2
10
0")
    End Sub
End Module
