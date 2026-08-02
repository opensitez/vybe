' vybe-test: vb/vb_weak_reference_gc_collect/test_vb_gc_keep_alive_prevents_premature_collection
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

Class ResourceTracker
    Public Id As Integer = 100
End Class

Module Program
    Sub Main()
        Dim res As New ResourceTracker()
        Dim id = res.Id
        GC.KeepAlive(res) ' Ensures res is not collected prior to this line!
        __Check(CStr(id), "100")
    End Sub
End Module
