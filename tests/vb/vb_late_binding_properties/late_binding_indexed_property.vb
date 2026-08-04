' vybe-test: vb/vb_late_binding_properties/late_binding_indexed_property
' origin: languages/vb/tests/vb/test_vb_late_binding_properties.rs

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

Class ItemCollection
    Private items As String() = {"A", "B", "C"}
    
    Default Public Property Item(index As Integer) As String
        Get
            Return items(index)
        End Get
        Set(value As String)
            items(index) = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim col As Object = New ItemCollection()
        
        ' Late bound indexed property via Default property (implicit)
        col(1) = "Z"
        
        ' Late bound indexed property (explicit)
        __P(CStr(col.Item(0)))
        __P(CStr(col(1)))
        __Check("A
Z")
    End Sub
End Module
