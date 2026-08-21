' vybe-test: vb/vb_index_out_of_range_exception/test_vb_indexed_property_out_of_bounds_custom_handling
' origin: languages/vb/tests/vb/test_vb_index_out_of_range_exception.rs

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

Imports System
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


Class SafeArray
    Private data(2) As Integer
    Default Public Property Item(idx As Integer) As Integer
        Get
            If idx < 0 OrElse idx >= data.Length Then
                Throw New IndexOutOfRangeException("SafeArray index out of bounds")
            End If
            Return data(idx)
        End Get
        Set(value As Integer)
            If idx < 0 OrElse idx >= data.Length Then
                Throw New IndexOutOfRangeException("SafeArray index out of bounds")
            End If
            data(idx) = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim sa As New SafeArray()
        Try
            sa(10) = 42
        Catch ex As IndexOutOfRangeException
            __P(CStr(ex.Message))
        End Try
        __Check("SafeArray index out of bounds")
    End Sub
End Module
