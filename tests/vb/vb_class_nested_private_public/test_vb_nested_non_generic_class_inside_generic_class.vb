' vybe-test: vb/vb_class_nested_private_public/test_vb_nested_non_generic_class_inside_generic_class
' origin: languages/vb/tests/vb/test_vb_class_nested_private_public.rs

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

Class OuterList(Of T)
    Public Class Node
        Public Element As T
        Public NextNode As Node
        Public Sub New(e As T)
            Element = e
        End Sub
    End Class
End Class

Module Program
    Sub Main()
        Dim node1 As New OuterList(Of Integer).Node(10)
        Dim node2 As New OuterList(Of Integer).Node(20)
        node1.NextNode = node2
        __P(CStr(node1.Element & "->" & node1.NextNode.Element))
        __Check("10->20")
    End Sub
End Module
