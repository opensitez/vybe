' vybe-test: vb/vb_weak_reference_gc_collect/test_vb_weak_reference_set_target_reassignment
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

Class Token
    Public Name As String
    Public Sub New(n As String)
        Name = n
    End Sub
End Class

Module Program
    Sub Main()
        Dim t1 As New Token("T1")
        Dim t2 As New Token("T2")
        Dim weakRef As New WeakReference(Of Token)(t1)

        weakRef.SetTarget(t2)
        Dim target As Token = Nothing
        weakRef.TryGetTarget(target)
        __Check(CStr(target.Name), "T2")
    End Sub
End Module
