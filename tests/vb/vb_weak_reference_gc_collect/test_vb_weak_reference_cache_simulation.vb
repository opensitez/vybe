' vybe-test: vb/vb_weak_reference_gc_collect/test_vb_weak_reference_cache_simulation
' origin: languages/vb/tests/vb/test_vb_weak_reference_gc_collect.rs

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

Imports System
Imports System.Collections.Generic

Class CacheManager
    Private cache As New Dictionary(Of String, WeakReference(Of String))()

    Public Sub Add(key As String, val As String)
        cache(key) = New WeakReference(Of String)(val)
    End Sub

    Public Function GetVal(key As String) As String
        Dim weakRef As WeakReference(Of String) = Nothing
        If cache.TryGetValue(key, weakRef) Then
            Dim target As String = Nothing
            If weakRef.TryGetTarget(target) Then Return target
        End If
        Return Nothing
    End Function
End Class

Module Program
    Sub Main()
        Dim cm As New CacheManager()
        Dim item = "CachedValue"
        cm.Add("K1", item)
        __P(CStr(cm.GetVal("K1")))
        __Check("CachedValue")
    End Sub
End Module
