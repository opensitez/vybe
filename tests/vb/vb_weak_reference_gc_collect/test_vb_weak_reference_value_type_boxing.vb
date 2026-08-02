' vybe-test: vb/vb_weak_reference_gc_collect/test_vb_weak_reference_value_type_boxing
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

Module Program
    Sub Main()
        ' Boxed value type target in WeakReference
        Dim boxed As Object = 999
        Dim weakRef As New WeakReference(boxed)
        __Check(CStr(weakRef.Target.ToString()), "999")
    End Sub
End Module
