' vybe-test: vb/vb_interface_hierarchy_multiple_inheritance/test_vb_interface_indexer_multiple_overloads
' origin: languages/vb/tests/vb/test_vb_interface_hierarchy_multiple_inheritance.rs

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

Interface IIndexable
    Default Property Item(key As String) As String
    Default Property Item(index As Integer) As String
End Interface

Class DictionaryAdapter
    Implements IIndexable
    Public Property Item(key As String) As String Implements IIndexable.Item
        Get
            Return "Key_" & key
        End Get
        Set(value As String)
        End Set
    End Property
    Public Property Item(index As Integer) As String Implements IIndexable.Item
        Get
            Return "Idx_" & index
        End Get
        Set(value As String)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim idx As IIndexable = New DictionaryAdapter()
        __P(CStr(idx("name") & "|" & idx(42)))
        __Check("Key_name|Idx_42")
    End Sub
End Module
