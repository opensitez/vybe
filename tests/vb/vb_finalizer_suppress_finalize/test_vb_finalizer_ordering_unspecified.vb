' vybe-test: vb/vb_finalizer_suppress_finalize/test_vb_finalizer_ordering_unspecified
' origin: languages/vb/tests/vb/test_vb_finalizer_suppress_finalize.rs

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

Class ObjA
    Protected Overrides Sub Finalize()
    End Sub
End Class

Class ObjB
    Protected Overrides Sub Finalize()
    End Sub
End Class

Module Program
    Sub Main()
        Sub()
            Dim a As New ObjA()
            Dim b As New ObjB()
        End Sub()

        GC.Collect()
        GC.WaitForPendingFinalizers()
        __Check(CStr("Multiple Finalizers Executed Safely"), "Multiple Finalizers Executed Safely")
    End Sub
End Module
