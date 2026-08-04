' vybe-test: vb/vb_indexer_multidimensional/test_vb_indexer_overloaded_string_and_int
' origin: languages/vb/tests/vb/test_vb_indexer_multidimensional.rs

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

Class DataStore
    Private _byInt As New System.Collections.Generic.Dictionary(Of Integer, String)()
    Private _byStr As New System.Collections.Generic.Dictionary(Of String, String)()

    Default Public Property Item(id As Integer) As String
        Get
            Return _byInt(id)
        End Get
        Set(value As String)
            _byInt(id) = value
        End Set
    End Property

    Default Public Property Item(key As String) As String
        Get
            Return _byStr(key)
        End Get
        Set(value As String)
            _byStr(key) = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim ds As New DataStore()
        ds(1) = "NumOne"
        ds("A") = "StrA"
        __P(CStr(ds(1)))
        __P(CStr(ds("A")))
        __Check("NumOne
StrA")
    End Sub
End Module
