' vybe-test: vb/vb_weak_reference_gc_collect/test_vb_weak_reference_gc_collect_clears_target
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

Class DisposableTarget
End Class

Module Program
    Sub Main()
        Dim weakRef As WeakReference(Of DisposableTarget)
        Sub()
            Dim obj As New DisposableTarget()
            weakRef = New WeakReference(Of DisposableTarget)(obj)
        End Sub()

        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()

        Dim target As DisposableTarget = Nothing
        Dim isAlive = weakRef.TryGetTarget(target)
        __Check(CStr(isAlive), "False")
    End Sub
End Module
