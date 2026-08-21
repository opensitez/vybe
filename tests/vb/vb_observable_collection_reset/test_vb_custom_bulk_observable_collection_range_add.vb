' vybe-test: vb/vb_observable_collection_reset/test_vb_custom_bulk_observable_collection_range_add
' origin: languages/vb/tests/vb/test_vb_observable_collection_reset.rs

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
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
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


Class RangeObservableCollection(Of T)
    Inherits ObservableCollection(Of T)

    Public Sub AddRange(items As IEnumerable(Of T))
        For Each item In items
            Items.Add(item)
        Next
        OnCollectionChanged(New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset))
    End Sub
End Class

Module Program
    Sub Main()
        Dim col As New RangeObservableCollection(Of Integer)()
        Dim resetCount = 0
        AddHandler col.CollectionChanged, Sub(s, e)
            If e.Action = NotifyCollectionChangedAction.Reset Then resetCount += 1
            __Check("3|Reset=1")
        End Sub
        col.AddRange({10, 20, 30})
        __P(CStr(col.Count & "|Reset=" & resetCount))
    End Sub
End Module
