' vybe-test: vb/vb_finalizer_suppress_finalize/test_vb_gc_re_register_for_finalize
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

Class ReRegisteredObject
    Public Shared Executions As Integer = 0

    Protected Overrides Sub Finalize()
        Executions += 1
    End Sub

    Public Sub ReRegister()
        GC.ReRegisterForFinalize(Me)
    End Sub
End Class

Module Program
    Sub Main()
        Dim obj As New ReRegisteredObject()
        GC.SuppressFinalize(obj)
        obj.ReRegister() ' Re-enable finalization!

        obj = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()

        __Check(CStr("Finalized Count: " & ReRegisteredObject.Executions), "Finalized Count: 1")
    End Sub
End Module
