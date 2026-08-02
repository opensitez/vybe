' vybe-test: vb/vb_exception_stack_trace_preservation/test_vb_exception_in_destructor_finalizer_suppressed
' origin: languages/vb/tests/vb/test_vb_exception_stack_trace_preservation.rs

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

Class Destructible
    Protected Overrides Sub Finalize()
        Try
            ' Destructors should never throw uncaught exceptions
            __Check(CStr("Finalizer executed"), "Finalizer executed")
        Catch ex As Exception
        Finally
            MyBase.Finalize()
        End Try
    End Sub
End Class

Module Program
    Sub Main()
        Dim d As New Destructible()
        d = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub
End Module
