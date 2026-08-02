' vybe-test: vb/vb_finalizer_suppress_finalize/test_vb_finalizer_re_register_null_throws
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

Module Program
    Sub Main()
        Try
            GC.ReRegisterForFinalize(Nothing)
        Catch ex As ArgumentNullException
            __Check(CStr("ArgumentNullException Caught on Null ReRegisterForFinalize"), "ArgumentNullException Caught on Null ReRegisterForFinalize")
        End Try
    End Sub
End Module
