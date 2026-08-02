' vybe-test: vb/vb_try_catch_rethrow_throw/test_vb_throw_null_reference_throws_null_pointer
' origin: languages/vb/tests/vb/test_vb_try_catch_rethrow_throw.rs

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
            Dim nullEx As Exception = Nothing
            Throw nullEx
        Catch ex As NullReferenceException
            __Check(CStr("Caught NullReferenceException on throw null"), "Caught NullReferenceException on throw null")
        End Try
    End Sub
End Module
