' vybe-test: vb/vb_gc_add_memory_pressure/test_vb_gc_get_generation_from_weak_reference
' origin: languages/vb/tests/vb/test_vb_gc_add_memory_pressure.rs

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

Class Target
End Class

Module Program
    Sub Main()
        Dim obj As New Target()
        Dim weakRef As New WeakReference(obj)
        Dim gen = GC.GetGeneration(weakRef)
        __Check(CStr(gen >= 0), "True")
    End Sub
End Module
