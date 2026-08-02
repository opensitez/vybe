' vybe-test: vb/vb_try_catch_rethrow_throw/test_vb_try_catch_rethrow_wrapped_exception
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
    Private Sub Process()
        Try
            Throw New FormatException("Invalid Number")
        Catch ex As FormatException
            Throw New InvalidOperationException("Process Failed", ex)
        End Try
    End Sub

    Sub Main()
        Try
            Process()
        Catch ex As InvalidOperationException
            __Check(CStr(ex.Message & " | Inner: " & ex.InnerException.Message), "Process Failed | Inner: Invalid Number")
        End Try
    End Sub
End Module
