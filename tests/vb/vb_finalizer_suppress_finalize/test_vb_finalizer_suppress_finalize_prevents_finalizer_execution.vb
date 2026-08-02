' vybe-test: vb/vb_finalizer_suppress_finalize/test_vb_finalizer_suppress_finalize_prevents_finalizer_execution
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

Class SuppressedFinalizerObject
    Public Shared FinalizerRan As Boolean = False

    Protected Overrides Sub Finalize()
        FinalizerRan = True
    End Sub

    Public Sub Cleanup()
        GC.SuppressFinalize(Me)
    End Sub
End Class

Module Program
    Sub Main()
        Sub()
            Dim obj As New SuppressedFinalizerObject()
            obj.Cleanup()
        End Sub()

        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()

        __Check(CStr(SuppressedFinalizerObject.FinalizerRan), "False")
    End Sub
End Module
