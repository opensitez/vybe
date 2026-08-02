' vybe-test: vb/vb_finally_block_execution_guarantees/test_vb_finally_executes_on_return
' origin: languages/vb/tests/vb/test_vb_finally_block_execution_guarantees.rs

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
    Function TestFunc() As String
        Try
            Return "ReturnedValue"
        Finally
            __Check(CStr("FinallyExecuted"), "FinallyExecuted")
        End Try
    End Function

    Sub Main()
        Dim res As String = TestFunc()
        __Check(CStr(res), "ReturnedValue")
    End Sub
End Module
