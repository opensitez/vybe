' vybe-test: vb/vb_try_catch_rethrow_throw/test_vb_exception_data_dictionary
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
            Dim ex As New Exception("Data Error")
            ex.Data("TimeStamp") = "2025-01-01"
            Throw ex
        Catch ex As Exception
            __Check(CStr("TimeStamp: " & ex.Data("TimeStamp").ToString()), "TimeStamp: 2025-01-01")
        End Try
    End Sub
End Module
