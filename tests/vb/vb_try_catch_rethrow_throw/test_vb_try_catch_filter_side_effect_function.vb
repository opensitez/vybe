' vybe-test: vb/vb_try_catch_rethrow_throw/test_vb_try_catch_filter_side_effect_function
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
    Private Function LogAndCheck(ex As Exception) As Boolean
        __Check(CStr("Filter Evaluated for: " & ex.Message), "Filter Evaluated for: TestMsg")
        Return True
    End Function

    Sub Main()
        Try
            Throw New InvalidOperationException("TestMsg")
        Catch ex As Exception When LogAndCheck(ex)
            __Check(CStr("Catch Executed"), "Catch Executed")
        End Try
    End Sub
End Module
