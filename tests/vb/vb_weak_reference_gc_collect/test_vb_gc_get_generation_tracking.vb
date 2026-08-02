' vybe-test: vb/vb_weak_reference_gc_collect/test_vb_gc_get_generation_tracking
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

Class PersistentObj
End Class

Module Program
    Sub Main()
        Dim obj As New PersistentObj()
        Dim gen0 = GC.GetGeneration(obj)
        __Check(CStr(gen0 >= 0), "True")
    End Sub
End Module
