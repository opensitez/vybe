' vybe-test: vb/vb_finalizer_suppress_finalize/test_vb_finalizer_swallows_unhandled_exception
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

Class FaultyFinalizer
    Protected Overrides Sub Finalize()
        ' Exceptions in finalizer are swallowed by CLR runtime without crashing application in default policy
    End Sub
End Class

Module Program
    Sub Main()
        Sub()
            Dim f As New FaultyFinalizer()
        End Sub()
        GC.Collect()
        GC.WaitForPendingFinalizers()
        __Check(CStr("Completed Safe GC"), "Completed Safe GC")
    End Sub
End Module
