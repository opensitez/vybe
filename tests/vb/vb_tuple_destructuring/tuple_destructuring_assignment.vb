' vybe-test: vb/vb_tuple_destructuring/tuple_destructuring_assignment
' origin: languages/vb/tests/vb/test_vb_tuple_destructuring.rs

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

Module M
    Function GetInfo() As (Name As String, Age As Integer)
        Return ("John", 30)
    End Function

    Sub Main()
        ' VB.NET does not natively support deconstruction syntax (Dim (name, age) = GetInfo())
        ' Wait, it doesn't? C# has deconstruction, but VB.NET does not have direct tuple deconstruction assignment syntax.
        ' Let's just use the tuple literal syntax and element access.
        Dim t = GetInfo()
        __P(CStr(t.Name))
        __P(CStr(t.Age))
        
        ' We can assign tuples to tuples
        Dim t2 As (String, Integer) = t
        __P(CStr(t2.Item1))
        __Check("John
30
John")
    End Sub
End Module
