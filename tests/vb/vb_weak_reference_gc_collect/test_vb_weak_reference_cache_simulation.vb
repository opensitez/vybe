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

Module VybeCheck
    Sub __Check(got As String, want As String)
        If got <> want Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & got & "]")
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
        __Check(CStr(cm.GetVal("K1")), "CachedValue")
    End Sub
End Module
