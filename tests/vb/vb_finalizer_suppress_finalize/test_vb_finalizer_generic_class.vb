' vybe-test: vb/vb_finalizer_suppress_finalize/test_vb_finalizer_generic_class
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

Class GenericFinalizer(Of T)
    Public Shared Finalized As Boolean = False
    Protected Overrides Sub Finalize()
        Finalized = True
    End Sub
End Class

Module Program
    Sub Main()
        Sub()
            Dim g As New GenericFinalizer(Of String)()
        End Sub()

        GC.Collect()
        GC.WaitForPendingFinalizers()

        __Check(CStr(GenericFinalizer(Of String).Finalized), "True")
    End Sub
End Module
