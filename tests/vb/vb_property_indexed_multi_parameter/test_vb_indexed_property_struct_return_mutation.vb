' vybe-test: vb/vb_property_indexed_multi_parameter/test_vb_indexed_property_struct_return_mutation
' origin: languages/vb/tests/vb/test_vb_property_indexed_multi_parameter.rs

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

Structure Point
    Public X As Integer
    Public Y As Integer
End Structure

Class PointGrid
    Private points(1) As Point
    Default Public Property Item(idx As Integer) As Point
        Get
            Return points(idx)
        End Get
        Set(value As Point)
            points(idx) = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim pg As New PointGrid()
        pg(0) = New Point With {.X = 10, .Y = 20}
        __P(CStr(pg(0).X & "," & pg(0).Y))
        __Check("10,20")
    End Sub
End Module
