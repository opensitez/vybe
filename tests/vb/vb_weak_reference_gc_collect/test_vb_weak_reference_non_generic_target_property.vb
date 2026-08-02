' vybe-test: vb/vb_weak_reference_gc_collect/test_vb_weak_reference_non_generic_target_property
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

Class Sample
    Public Val As Integer = 42
End Class

Module Program
    Sub Main()
        Dim obj As New Sample()
        Dim weakRef As New WeakReference(obj)
        Dim target As Sample = CType(weakRef.Target, Sample)
        __Check(CStr(weakRef.IsAlive & "|" & target.Val), "True|42")
    End Sub
End Module
